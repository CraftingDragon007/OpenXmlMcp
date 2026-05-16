using System.Collections.Concurrent;
using System.Text.Json;
using OpenXmlMcp.Server.Models;

namespace OpenXmlMcp.Server.Services;

/// <summary>
/// Manages session lifecycle, per-session write locks, snapshot-based undo, and operation logs.
/// All document-type-specific services depend on this for session access and write orchestration.
/// </summary>
public class SessionManager
{
    private const long MaxOpenFileSizeBytes = 20 * 1024 * 1024;

    private readonly ConcurrentDictionary<string, OfficeSession> _sessions = new();
    private readonly ConcurrentDictionary<string, Stack<string>> _sessionSnapshots = new();
    private readonly ConcurrentDictionary<string, List<OperationLogEntry>> _sessionOperationLog = new();
    private readonly ConcurrentDictionary<string, object> _sessionWriteLocks = new();

    public string OpenDocument(string filePath, bool readOnly = false)
    {
        var normalizedPath = NormalizePath(filePath);

        if (!File.Exists(normalizedPath))
        {
            throw new FileNotFoundException("Office file not found.", normalizedPath);
        }

        var fileInfo = new FileInfo(normalizedPath);
        if (fileInfo.Length > MaxOpenFileSizeBytes)
        {
            throw new InvalidOperationException($"File size {fileInfo.Length} exceeds safety limit {MaxOpenFileSizeBytes} bytes.");
        }

        var session = new OfficeSession
        {
            Id = Guid.NewGuid().ToString("N"),
            FilePath = normalizedPath,
            DocumentType = GetDocumentTypeFromPath(normalizedPath),
            IsReadOnly = readOnly,
            CreatedAtUtc = DateTimeOffset.UtcNow
        };

        _sessions[session.Id] = session;
        _sessionSnapshots.TryAdd(session.Id, new Stack<string>());
        _sessionOperationLog.TryAdd(session.Id, new List<OperationLogEntry>());
        _sessionWriteLocks.TryAdd(session.Id, new object());
        AppendOperationLog(session.Id, "open_document", $"Opened '{normalizedPath}' (readOnly={readOnly}).");
        return session.Id;
    }

    public string CloseDocument(string sessionId)
    {
        if (_sessionSnapshots.TryRemove(sessionId, out var snapshots))
        {
            foreach (var snapshotFile in snapshots)
            {
                DeleteIfExists(snapshotFile);
            }
        }

        _sessionOperationLog.TryRemove(sessionId, out _);
        _sessionWriteLocks.TryRemove(sessionId, out _);

        if (!_sessions.TryRemove(sessionId, out _))
        {
            throw new InvalidOperationException($"Session '{sessionId}' was not found.");
        }

        return BuildMutationResult("close_document", new { sessionId });
    }

    public string UndoLastChange(string sessionId)
    {
        var session = GetSession(sessionId);
        if (session.IsReadOnly)
        {
            throw new InvalidOperationException("Session is read-only.");
        }

        if (!_sessionSnapshots.TryGetValue(sessionId, out var snapshots) || snapshots.Count == 0)
        {
            throw new InvalidOperationException("No checkpoint is available to undo.");
        }

        var snapshotPath = snapshots.Pop();
        File.Copy(snapshotPath, session.FilePath, overwrite: true);
        DeleteIfExists(snapshotPath);
        AppendOperationLog(sessionId, "undo_last_change", "Restored document from latest checkpoint.");
        return BuildMutationResult("undo_last_change", new { sessionId });
    }

    public string GetOperationHistory(string sessionId)
    {
        _ = GetSession(sessionId);
        var history = _sessionOperationLog.TryGetValue(sessionId, out var entries)
            ? entries.Select(x => new { x.TimestampUtc, x.CanonicalOperationName, x.PublicOperationName, x.OperationName, x.Message }).ToArray()
            : [];

        return JsonSerializer.Serialize(new { sessionId, count = history.Length, history });
    }

    public OfficeSession GetSession(string sessionId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sessionId);
        if (_sessions.TryGetValue(sessionId, out var session))
        {
            return session;
        }

        throw new InvalidOperationException($"Session '{sessionId}' was not found.");
    }

    public OfficeSession GetWritableSession(string sessionId, OfficeDocumentType requiredType)
    {
        var session = GetSession(sessionId);
        if (session.IsReadOnly)
        {
            throw new InvalidOperationException("Session is read-only.");
        }

        if (session.DocumentType != requiredType)
        {
            throw new InvalidOperationException($"Session document type must be '{requiredType}'.");
        }

        return session;
    }

    public void ExecuteWriteOperation(OfficeSession session, string operationName, Action operation, string? publicOperationName = null)
    {
        var writeLock = _sessionWriteLocks.GetOrAdd(session.Id, _ => new object());
        lock (writeLock)
        {
            var snapshotPath = CreateSnapshot(session);

            try
            {
                operation();
                AppendOperationLog(session.Id, operationName, "Operation completed.", publicOperationName);
            }
            catch
            {
                if (_sessionSnapshots.TryGetValue(session.Id, out var snapshots) && snapshots.Count > 0 && snapshots.Peek() == snapshotPath)
                {
                    snapshots.Pop();
                }

                DeleteIfExists(snapshotPath);
                throw;
            }
        }
    }

    public void AppendOperationLog(string sessionId, string operationName, string message, string? publicOperationName = null)
    {
        var history = _sessionOperationLog.GetOrAdd(sessionId, _ => new List<OperationLogEntry>());
        history.Add(new OperationLogEntry(DateTimeOffset.UtcNow, operationName, publicOperationName ?? operationName, message));
    }

    private string CreateSnapshot(OfficeSession session)
    {
        var snapshotPath = Path.Combine(Path.GetTempPath(), $"openxmlmcp-snapshot-{session.Id}-{Guid.NewGuid():N}{Path.GetExtension(session.FilePath)}");
        File.Copy(session.FilePath, snapshotPath, overwrite: true);

        var snapshots = _sessionSnapshots.GetOrAdd(session.Id, _ => new Stack<string>());
        snapshots.Push(snapshotPath);
        AppendOperationLog(session.Id, "checkpoint", $"Snapshot created: {snapshotPath}");
        return snapshotPath;
    }

    private static void DeleteIfExists(string filePath)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }

    private static string NormalizePath(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        return Path.GetFullPath(filePath);
    }

    private static OfficeDocumentType GetDocumentTypeFromPath(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension switch
        {
            ".docx" => OfficeDocumentType.Word,
            ".xlsx" => OfficeDocumentType.Excel,
            ".pptx" => OfficeDocumentType.PowerPoint,
            _ => throw new InvalidOperationException($"Unsupported extension '{extension}'.")
        };
    }

    public static string BuildMutationResult(string operation, object target, bool changed = true, string? publicOperationName = null, string? canonicalOperationName = null)
    {
        return JsonSerializer.Serialize(new
        {
            ok = true,
            changed,
            operation,
            publicOperationName = publicOperationName ?? operation,
            canonicalOperationName = canonicalOperationName ?? operation,
            target
        });
    }

    private sealed record OperationLogEntry(DateTimeOffset TimestampUtc, string CanonicalOperationName, string PublicOperationName, string Message)
    {
        public string OperationName => CanonicalOperationName;
    }
}
