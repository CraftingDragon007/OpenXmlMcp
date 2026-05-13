using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlMcp.Server.Models;

namespace OpenXmlMcp.Server.Services;

public class OfficeSessionService
{
    private const int DefaultParagraphBeforePt = 0;
    private const int DefaultParagraphAfterPt = 8;
    private const double DefaultParagraphLineSpacing = 1.15;
    private const int DefaultHeadingBeforePt = 12;
    private const int DefaultHeadingAfterPt = 6;
    private const double DefaultHeadingLineSpacing = 1.15;
    private const int DefaultListBeforePt = 0;
    private const int DefaultListAfterPt = 2;
    private const double DefaultListLineSpacing = 1.15;
    private readonly ConcurrentDictionary<string, OfficeSession> _sessions = new();
    private readonly ConcurrentDictionary<string, Stack<string>> _sessionSnapshots = new();
    private readonly ConcurrentDictionary<string, List<OperationLogEntry>> _sessionOperationLog = new();
    private const long MaxOpenFileSizeBytes = 20 * 1024 * 1024;

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
        AppendOperationLog(session.Id, "open_document", $"Opened '{normalizedPath}' (readOnly={readOnly}).");
        return session.Id;
    }

    public string CreateDocument(string filePath, string documentType)
    {
        var normalizedPath = NormalizePath(filePath);
        Directory.CreateDirectory(Path.GetDirectoryName(normalizedPath) ?? Directory.GetCurrentDirectory());

        var type = ParseDocumentType(documentType);
        CreateEmptyDocument(normalizedPath, type);

        return OpenDocument(normalizedPath);
    }

    public void SaveDocument(string sessionId)
    {
        _ = GetSession(sessionId);
    }

    public void CloseDocument(string sessionId)
    {
        if (_sessionSnapshots.TryRemove(sessionId, out var snapshots))
        {
            foreach (var snapshotFile in snapshots)
            {
                DeleteIfExists(snapshotFile);
            }
        }

        _sessionOperationLog.TryRemove(sessionId, out _);

        if (!_sessions.TryRemove(sessionId, out _))
        {
            throw new InvalidOperationException($"Session '{sessionId}' was not found.");
        }
    }

    public string GetOperationHistory(string sessionId)
    {
        _ = GetSession(sessionId);
        var history = _sessionOperationLog.TryGetValue(sessionId, out var entries)
            ? entries.Select(x => new { x.TimestampUtc, x.OperationName, x.Message }).ToArray()
            : [];

        return JsonSerializer.Serialize(new { sessionId, count = history.Length, history });
    }

    public void UndoLastChange(string sessionId)
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
    }

    public string GetDocumentInfo(string sessionId)
    {
        var session = GetSession(sessionId);
        var fileInfo = new FileInfo(session.FilePath);

        var payload = new
        {
            session.Id,
            session.FilePath,
            documentType = session.DocumentType.ToString(),
            session.IsReadOnly,
            session.CreatedAtUtc,
            exists = fileInfo.Exists,
            sizeBytes = fileInfo.Exists ? fileInfo.Length : 0
        };

        return JsonSerializer.Serialize(payload);
    }

    public string ListStructure(string sessionId)
    {
        var session = GetSession(sessionId);

        object payload = session.DocumentType switch
        {
            OfficeDocumentType.Word => ListWordStructure(session.FilePath),
            OfficeDocumentType.Excel => ListExcelStructure(session.FilePath),
            OfficeDocumentType.PowerPoint => ListPowerPointStructure(session.FilePath),
            _ => throw new InvalidOperationException("Unsupported document type.")
        };

        return JsonSerializer.Serialize(payload);
    }

    public string FindText(string sessionId, string query)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(query);
        var session = GetSession(sessionId);

        object payload = session.DocumentType switch
        {
            OfficeDocumentType.Word => FindTextInWord(session.FilePath, query),
            OfficeDocumentType.Excel => FindTextInExcel(session.FilePath, query),
            OfficeDocumentType.PowerPoint => FindTextInPowerPoint(session.FilePath, query),
            _ => throw new InvalidOperationException("Unsupported document type.")
        };

        return JsonSerializer.Serialize(payload);
    }

    public string ValidateOperation(string sessionId, string operationName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(operationName);
        var session = GetSession(sessionId);

        var normalized = NormalizeOperationName(operationName);
        var spec = TryGetOperationSpec(normalized);

        var warnings = new List<string>();
        var isValid = true;

        if (spec is null)
        {
            isValid = false;
            warnings.Add($"Unknown operation '{operationName}'.");
        }

        if (spec?.RequiresWrite == true && session.IsReadOnly)
        {
            isValid = false;
            warnings.Add("Session is read-only.");
        }

        if (spec?.ExpectedType is not null && session.DocumentType != spec.ExpectedType.Value)
        {
            isValid = false;
            warnings.Add($"Operation '{operationName}' expects '{spec.ExpectedType.Value}' document type.");
        }

        var payload = new
        {
            sessionId,
            operationName,
            isValid,
            warnings
        };

        return JsonSerializer.Serialize(payload);
    }

    public void WordAddTable(string sessionId, int rows, int columns)
    {
        if (rows <= 0 || columns <= 0)
        {
            throw new InvalidOperationException("Rows and columns must be greater than zero.");
        }

        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_add_table", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");

            var table = new W.Table();
            for (var r = 0; r < rows; r++)
            {
                var tableRow = new W.TableRow();
                for (var c = 0; c < columns; c++)
                {
                    tableRow.Append(new W.TableCell(new W.Paragraph(new W.Run(new W.Text(string.Empty)))));
                }

                table.Append(tableRow);
            }

            body.Append(table);
            document.MainDocumentPart.Document?.Save();
        });
    }

    public void WordInsertParagraphAt(string sessionId, int index, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);

        ExecuteWriteOperation(session, "word_insert_paragraph_at", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();

            if (index < 1 || index > paragraphs.Count + 1)
            {
                throw new InvalidOperationException($"Index {index} is out of range. Valid range is 1..{paragraphs.Count + 1}.");
            }

            var newParagraph = new W.Paragraph(new W.Run(new W.Text(text)));
            ApplyWordParagraphSpacing(newParagraph, DefaultParagraphBeforePt, DefaultParagraphAfterPt, DefaultParagraphLineSpacing);
            if (paragraphs.Count == 0 || index == paragraphs.Count + 1)
            {
                body.Append(newParagraph);
            }
            else
            {
                paragraphs[index - 1].InsertBeforeSelf(newParagraph);
            }

            document.MainDocumentPart.Document?.Save();
        });
    }

    public int WordReplaceText(string sessionId, string find, string replace, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(find);
        replace ??= string.Empty;
        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        var replacements = 0;

        ExecuteWriteOperation(session, "word_replace_text", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var textNodes = document.MainDocumentPart?.Document?.Descendants<W.Text>()
                ?? throw new InvalidOperationException("Word document text nodes are missing.");

            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            foreach (var textNode in textNodes)
            {
                if (string.IsNullOrEmpty(textNode.Text))
                {
                    continue;
                }

                var count = CountOccurrences(textNode.Text, find, comparison);
                if (count == 0)
                {
                    continue;
                }

                textNode.Text = ReplaceWithComparison(textNode.Text, find, replace, comparison);
                replacements += count;
            }

            document.MainDocumentPart.Document?.Save();
        });

        return replacements;
    }

    public void WordAddHeading(string sessionId, int level, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        if (level < 1 || level > 6)
        {
            throw new InvalidOperationException("Heading level must be between 1 and 6.");
        }

        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_add_heading", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");

            var heading = new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = $"Heading{level}" }),
                new W.Run(new W.Text(text)));
            ApplyWordParagraphSpacing(heading, DefaultHeadingBeforePt, DefaultHeadingAfterPt, DefaultHeadingLineSpacing);

            body.Append(heading);
            document.MainDocumentPart.Document?.Save();
        });
    }

    public void WordAddBulletedList(string sessionId, string lines, string bulletStyle = "disc")
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(lines);
        var items = lines.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(x => new WordListItem(x, 0, "bulleted", BulletStyle: bulletStyle))
            .ToArray();
        WordAddStructuredList(sessionId, JsonSerializer.Serialize(items));
    }

    public void WordAddNumberedList(string sessionId, string lines, string numberStyle = "decimal-dot")
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(lines);
        var items = lines.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(x => new WordListItem(x, 0, "numbered", NumberStyle: numberStyle))
            .ToArray();
        WordAddStructuredList(sessionId, JsonSerializer.Serialize(items));
    }

    public void WordAddStructuredList(string sessionId, string itemsJson)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(itemsJson);
        var items = JsonSerializer.Deserialize<List<WordListItem>>(itemsJson, new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("Invalid itemsJson for structured list.");
        if (items.Count == 0)
        {
            throw new InvalidOperationException("At least one list item is required.");
        }

        if (items.Any(i => string.IsNullOrWhiteSpace(i.Text)))
        {
            throw new InvalidOperationException("List items must contain text.");
        }

        if (items.Any(i => i.Level < 0 || i.Level > 8))
        {
            throw new InvalidOperationException("List levels must be between 0 and 8.");
        }

        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_add_structured_list", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var definitions = EnsureStructuredListNumbering(document, items);

            foreach (var item in items)
            {
                var normalizedKind = NormalizeListKind(item.Kind);
                var numberingId = normalizedKind == "numbered" ? definitions.NumberedNumberingId : definitions.BulletedNumberingId;
                var paragraph = new W.Paragraph(
                    new W.ParagraphProperties(
                        new W.NumberingProperties(
                            new W.NumberingLevelReference { Val = item.Level },
                            new W.NumberingId { Val = numberingId })),
                    new W.Run(new W.Text(item.Text)));
                ApplyWordParagraphSpacing(paragraph, DefaultListBeforePt, DefaultListAfterPt, DefaultListLineSpacing);
                body.Append(paragraph);
            }

            document.MainDocumentPart.Document?.Save();
        });
    }

    public void ExcelAddWorksheet(string sessionId, string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Excel);

        ExecuteWriteOperation(session, "excel_add_worksheet", () =>
        {
            using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, true);
            var workbookPart = GetWorkbookPart(spreadsheet);
            var workbook = GetWorkbook(workbookPart);
            var sheets = workbook.Sheets ?? workbook.AppendChild(new Sheets());

            if (GetSheets(workbookPart).Any(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException($"Sheet '{sheetName}' already exists.");
            }

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            worksheetPart.Worksheet.Save();

            var nextSheetId = sheets.Elements<Sheet>().Select(s => s.SheetId?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1;
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = nextSheetId,
                Name = sheetName
            });

            workbook.Save();
        });
    }

    public void PowerPointAddBulletSlide(string sessionId, string title, string bulletLines)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(bulletLines);
        var body = string.Join(Environment.NewLine, bulletLines.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(line => $"- {line}"));
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        ExecuteWriteOperation(session, "powerpoint_add_bullet_slide", () => PowerPointAddSlideCore(session.FilePath, title, body));
    }

    public string BatchExecute(string sessionId, string operationsJson)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(operationsJson);
        _ = GetSession(sessionId);

        var root = JsonNode.Parse(operationsJson) ?? throw new InvalidOperationException("Invalid operations JSON.");
        var operations = root.AsArray();
        var executed = 0;
        var failures = new List<object>();

        for (var index = 0; index < operations.Count; index++)
        {
            var op = operations[index];
            if (op is null)
            {
                continue;
            }

            var operation = op["operation"]?.GetValue<string>() ?? string.Empty;
            try
            {
                ExecuteBatchOperation(sessionId, operation, op);
                executed++;
            }
            catch (Exception ex)
            {
                failures.Add(new { operation, index, errorCode = GetErrorCode(ex), error = ex.Message });
            }
        }

        return JsonSerializer.Serialize(new
        {
            total = operations.Count,
            executed,
            failed = failures.Count,
            failures
        });
    }

    public void ApplyStylePreset(string sessionId, string preset = "default")
    {
        var currentSession = GetSession(sessionId);
        var session = GetWritableSession(sessionId, currentSession.DocumentType);
        var normalizedPreset = string.IsNullOrWhiteSpace(preset) ? "default" : preset.Trim().ToLowerInvariant();

        ExecuteWriteOperation(session, "apply_style_preset", () =>
        {
            switch (session.DocumentType)
            {
                case OfficeDocumentType.Word:
                    ApplyWordStylePreset(session.FilePath);
                    break;
                case OfficeDocumentType.Excel:
                    ApplyExcelStylePreset(session.FilePath);
                    break;
                case OfficeDocumentType.PowerPoint:
                    ApplyPowerPointStylePreset(session.FilePath, normalizedPreset);
                    break;
                default:
                    throw new InvalidOperationException("Unsupported document type.");
            }
        });
    }

    public void ApplyTextPreset(string sessionId, string preset, int targetIndex = 1)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(preset);
        var session = GetSession(sessionId);
        var normalized = preset.Trim().ToLowerInvariant();

        switch (session.DocumentType)
        {
            case OfficeDocumentType.Word:
                WordSetParagraphStyle(sessionId, targetIndex, BuildTextStyleFromPreset(normalized, isTitleSlot: true));
                break;
            case OfficeDocumentType.Excel:
                ExcelSetCellStyle(sessionId, "Sheet1", "A1", BuildTextStyleFromPreset(normalized, isTitleSlot: true));
                break;
            case OfficeDocumentType.PowerPoint:
                PowerPointSetTextStyle(sessionId, targetIndex, 0, BuildTextStyleFromPreset(normalized, isTitleSlot: true));
                if (normalized is "title" or "subtitle")
                {
                    PowerPointSetTextStyle(sessionId, targetIndex, 1, BuildTextStyleFromPreset(normalized, isTitleSlot: false));
                }

                break;
            default:
                throw new InvalidOperationException("Unsupported document type.");
        }
    }

    public void WordSetParagraphSpacing(string sessionId, int paragraphIndex, int beforePt, int afterPt, double lineSpacing)
    {
        if (beforePt < 0 || afterPt < 0)
        {
            throw new InvalidOperationException("Paragraph spacing before/after must be >= 0.");
        }

        if (lineSpacing <= 0)
        {
            throw new InvalidOperationException("Line spacing must be greater than 0.");
        }

        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_set_paragraph_spacing", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();
            if (paragraphIndex < 1 || paragraphIndex > paragraphs.Count)
            {
                throw new InvalidOperationException($"Paragraph index {paragraphIndex} is out of range. Valid range is 1..{paragraphs.Count}.");
            }

            ApplyWordParagraphSpacing(paragraphs[paragraphIndex - 1], beforePt, afterPt, lineSpacing);
            document.MainDocumentPart.Document?.Save();
        });
    }

    public void WordSetDocumentSpacingPreset(string sessionId, string preset = "normal")
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(preset);
        var normalized = preset.Trim().ToLowerInvariant();
        var settings = normalized switch
        {
            "compact" => (beforePt: 0, afterPt: 4, lineSpacing: 1.0),
            "comfortable" => (beforePt: 0, afterPt: 10, lineSpacing: 1.3),
            _ => (beforePt: DefaultParagraphBeforePt, afterPt: DefaultParagraphAfterPt, lineSpacing: DefaultParagraphLineSpacing)
        };

        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_set_document_spacing_preset", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            foreach (var paragraph in body.Elements<W.Paragraph>())
            {
                ApplyWordParagraphSpacing(paragraph, settings.beforePt, settings.afterPt, settings.lineSpacing);
            }

            document.MainDocumentPart.Document?.Save();
        });
    }

    public int WordInsertParagraphAfterText(string sessionId, string anchorText, string text, int occurrence = 1, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(anchorText);
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        if (occurrence < 1)
        {
            throw new InvalidOperationException("Occurrence must be >= 1.");
        }

        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        var insertedIndex = -1;
        ExecuteWriteOperation(session, "word_insert_paragraph_after_text", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            var match = paragraphs
                .Select((p, i) => new { Paragraph = p, Index = i + 1 })
                .Where(x => x.Paragraph.InnerText.Contains(anchorText, comparison))
                .Skip(occurrence - 1)
                .FirstOrDefault()
                ?? throw new InvalidOperationException($"Anchor text '{anchorText}' occurrence {occurrence} was not found.");

            var newParagraph = new W.Paragraph(new W.Run(new W.Text(text)));
            ApplyWordParagraphSpacing(newParagraph, DefaultParagraphBeforePt, DefaultParagraphAfterPt, DefaultParagraphLineSpacing);
            match.Paragraph.InsertAfterSelf(newParagraph);
            insertedIndex = match.Index + 1;

            document.MainDocumentPart.Document?.Save();
        });

        return insertedIndex;
    }

    public bool WordInsertTextAfterText(string sessionId, string anchorText, string text, int occurrence = 1, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(anchorText);
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        if (occurrence < 1)
        {
            throw new InvalidOperationException("Occurrence must be >= 1.");
        }

        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        var replaced = false;
        ExecuteWriteOperation(session, "word_insert_text_after_text", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var textNodes = document.MainDocumentPart?.Document?.Descendants<W.Text>().ToList()
                ?? throw new InvalidOperationException("Word text nodes are missing.");

            var remaining = occurrence;
            foreach (var node in textNodes)
            {
                if (string.IsNullOrEmpty(node.Text))
                {
                    continue;
                }

                var index = node.Text.IndexOf(anchorText, comparison);
                if (index < 0)
                {
                    continue;
                }

                remaining--;
                if (remaining > 0)
                {
                    continue;
                }

                var insertPos = index + anchorText.Length;
                node.Text = node.Text[..insertPos] + text + node.Text[insertPos..];
                replaced = true;
                break;
            }

            if (!replaced)
            {
                throw new InvalidOperationException($"Anchor text '{anchorText}' occurrence {occurrence} was not found.");
            }

            document.MainDocumentPart?.Document?.Save();
        });

        return replaced;
    }

    public string WordListStyles(string sessionId)
    {
        var session = GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Word)
        {
            throw new InvalidOperationException("Session document type is not Word.");
        }

        using var document = WordprocessingDocument.Open(session.FilePath, false);
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<W.Style>() ?? [];
        var payload = styles.Select(style => new
        {
            styleId = style.StyleId?.Value ?? string.Empty,
            name = style.StyleName?.Val?.Value ?? string.Empty,
            type = style.Type?.Value.ToString() ?? string.Empty,
            isDefault = style.Default?.Value ?? false
        }).ToArray();

        return JsonSerializer.Serialize(new { count = payload.Length, styles = payload });
    }

    public void WordApplyStyleByName(string sessionId, int paragraphIndex, string styleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_apply_style_by_name", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();
            if (paragraphIndex < 1 || paragraphIndex > paragraphs.Count)
            {
                throw new InvalidOperationException($"Paragraph index {paragraphIndex} is out of range. Valid range is 1..{paragraphs.Count}.");
            }

            var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<W.Style>() ?? [];
            var style = styles.FirstOrDefault(s =>
                string.Equals(s.StyleId?.Value, styleName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(s.StyleName?.Val?.Value, styleName, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException($"Style '{styleName}' was not found in document.");

            var paragraph = paragraphs[paragraphIndex - 1];
            var paragraphProperties = paragraph.ParagraphProperties ??= new W.ParagraphProperties();
            paragraphProperties.ParagraphStyleId = new W.ParagraphStyleId { Val = style.StyleId?.Value };
            document.MainDocumentPart?.Document?.Save();
        });
    }

    public void WordCreateOrUpdateStyle(string sessionId, string styleName, string styleJson)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        ArgumentException.ThrowIfNullOrWhiteSpace(styleJson);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_create_or_update_style", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            var stylePart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
            stylePart.Styles ??= new W.Styles();

            var styleId = SanitizeStyleId(styleName);
            var style = stylePart.Styles.Elements<W.Style>().FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
            if (style is null)
            {
                style = new W.Style { Type = W.StyleValues.Paragraph, StyleId = styleId, CustomStyle = true };
                style.Append(new W.StyleName { Val = styleName });
                style.Append(new W.BasedOn { Val = "Normal" });
                style.Append(new W.NextParagraphStyle { Val = "Normal" });
                style.Append(new W.UIPriority { Val = 99 });
                style.Append(new W.UnhideWhenUsed());
                style.Append(new W.PrimaryStyle());
                stylePart.Styles.Append(style);
            }

            var options = JsonNode.Parse(styleJson) as JsonObject ?? throw new InvalidOperationException("Invalid styleJson object.");
            ApplyStyleOptions(style, options);
            stylePart.Styles.Save();
            mainPart.Document?.Save();
        });
    }

    public void WordSetParagraphStyle(string sessionId, int paragraphIndex, string fontName, int fontSize, bool bold, bool italic, string colorHex)
    {
        WordSetParagraphStyle(sessionId, paragraphIndex, new TextStyle(fontName, fontSize, bold, italic, colorHex));
    }

    public void ExcelSetCellStyle(string sessionId, string sheetName, string cellReference, string fontName, int fontSize, bool bold, bool italic, string colorHex)
    {
        ExcelSetCellStyle(sessionId, sheetName, cellReference, new TextStyle(fontName, fontSize, bold, italic, colorHex));
    }

    public void PowerPointSetTextStyle(string sessionId, int slideIndex, int slot, string fontName, int fontSize, bool bold, bool italic, string colorHex)
    {
        PowerPointSetTextStyle(sessionId, slideIndex, slot, new TextStyle(fontName, fontSize, bold, italic, colorHex));
    }

    private void WordSetParagraphStyle(string sessionId, int paragraphIndex, TextStyle style)
    {
        ValidateTextStyle(style);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_set_paragraph_style", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();

            if (paragraphIndex < 1 || paragraphIndex > paragraphs.Count)
            {
                throw new InvalidOperationException($"Paragraph index {paragraphIndex} is out of range. Valid range is 1..{paragraphs.Count}.");
            }

            ApplyWordParagraphStyle(paragraphs[paragraphIndex - 1], style);
            document.MainDocumentPart.Document?.Save();
        });
    }

    private void ExcelSetCellStyle(string sessionId, string sheetName, string cellReference, TextStyle style)
    {
        ValidateTextStyle(style);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);

        var session = GetWritableSession(sessionId, OfficeDocumentType.Excel);
        ExecuteWriteOperation(session, "excel_set_cell_style", () =>
        {
            using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, true);
            var workbookPart = GetWorkbookPart(spreadsheet);
            var sheet = GetSheetByName(workbookPart, sheetName);
            var worksheetPart = GetWorksheetPart(workbookPart, sheet);
            var worksheet = GetWorksheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            var rowIndex = ParseRowIndex(cellReference);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
            if (row is null)
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            var normalizedRef = cellReference.ToUpperInvariant();
            var cell = row.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, normalizedRef, StringComparison.OrdinalIgnoreCase));
            if (cell is null)
            {
                cell = new Cell { CellReference = normalizedRef, DataType = CellValues.String, CellValue = new CellValue(string.Empty) };
                row.Append(cell);
            }

            var styleIndex = EnsureExcelCellFormat(workbookPart, style);
            cell.StyleIndex = styleIndex;

            worksheet.Save();
            GetWorkbook(workbookPart).Save();
        });
    }

    private void PowerPointSetTextStyle(string sessionId, int slideIndex, int slot, TextStyle style)
    {
        ValidateTextStyle(style);
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);

        ExecuteWriteOperation(session, "powerpoint_set_text_style", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
            ApplyPowerPointSlotStyle(slidePart, slot, style);
            GetSlide(slidePart).Save();
        });
    }

    private static void ApplyWordStylePreset(string filePath)
    {
        using var document = WordprocessingDocument.Open(filePath, true);
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
        var stylePart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylePart.Styles ??= new W.Styles();

        if (!stylePart.Styles.Elements<W.DocDefaults>().Any())
        {
            stylePart.Styles.Append(new W.DocDefaults(
                new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
                    new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                    new W.FontSize { Val = "22" }))));
        }

        stylePart.Styles.Save();
    }

    private static void ApplyExcelStylePreset(string filePath)
    {
        using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet ??= new Stylesheet(
            new DocumentFormat.OpenXml.Spreadsheet.Fonts(
                new DocumentFormat.OpenXml.Spreadsheet.Font(
                    new FontName { Val = "Calibri" },
                    new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 11D }))
            { Count = 1U },
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
            { Count = 2U },
            new Borders(new DocumentFormat.OpenXml.Spreadsheet.Border()) { Count = 1U },
            new CellStyleFormats(new CellFormat()) { Count = 1U },
            new CellFormats(new CellFormat()) { Count = 1U });

        stylesPart.Stylesheet.Save();
        GetWorkbook(workbookPart).Save();
    }

    private static void ApplyPowerPointStylePreset(string filePath, string preset)
    {
        using var presentation = PresentationDocument.Open(filePath, true);
        var presentationPart = GetPresentationPart(presentation);
        _ = EnsurePresentationDefaults(presentationPart);

        var slideMasterPart = presentationPart.SlideMasterParts.First();
        var themePart = slideMasterPart.ThemePart ?? slideMasterPart.AddNewPart<ThemePart>();
        WriteDefaultTheme(themePart, ResolveThemeOptions(preset));

        slideMasterPart.SlideMaster?.Save();
        presentationPart.Presentation?.Save();
    }

    public void WordAppendParagraph(string sessionId, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_append_paragraph", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body;
            if (body is null)
            {
                throw new InvalidOperationException("Word document body is missing.");
            }

            body.AppendChild(new W.Paragraph(new W.Run(new W.Text(text))));
            var paragraph = body.Elements<W.Paragraph>().Last();
            ApplyWordParagraphSpacing(paragraph, DefaultParagraphBeforePt, DefaultParagraphAfterPt, DefaultParagraphLineSpacing);
            document.MainDocumentPart!.Document!.Save();
        });
    }

    public void ExcelSetCellValue(string sessionId, string sheetName, string cellReference, string value)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Excel);
        ExecuteWriteOperation(session, "excel_set_cell_value", () =>
        {
            using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, true);
            var workbookPart = GetWorkbookPart(spreadsheet);
            var sheet = GetSheetByName(workbookPart, sheetName);
            var worksheetPart = GetWorksheetPart(workbookPart, sheet);
            var worksheet = GetWorksheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            var rowIndex = ParseRowIndex(cellReference);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
            if (row is null)
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            var cell = row.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));
            if (cell is null)
            {
                cell = new Cell { CellReference = cellReference.ToUpperInvariant() };
                row.Append(cell);
            }

            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(value);

            worksheet.Save();
            GetWorkbook(workbookPart).Save();
        });
    }

    public string ExcelGetCellValue(string sessionId, string sheetName, string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Excel)
        {
            throw new InvalidOperationException("Session document type is not Excel.");
        }

        using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheet = GetSheetByName(workbookPart, sheetName);
        var worksheetPart = GetWorksheetPart(workbookPart, sheet);
        var worksheet = GetWorksheet(worksheetPart);
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));

        return cell?.CellValue?.Text ?? string.Empty;
    }

    public string ExcelGetUsedRange(string sessionId, string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        var session = GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Excel)
        {
            throw new InvalidOperationException("Session document type is not Excel.");
        }

        using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheet = GetSheetByName(workbookPart, sheetName);
        var worksheet = GetWorksheet(GetWorksheetPart(workbookPart, sheet));

        var cells = worksheet.Descendants<Cell>()
            .Where(c => !string.IsNullOrWhiteSpace(c.CellReference?.Value))
            .Select(c => ParseCellReference(c.CellReference!.Value!))
            .ToList();

        if (cells.Count == 0)
        {
            return JsonSerializer.Serialize(new { startCell = string.Empty, endCell = string.Empty, rowCount = 0, columnCount = 0 });
        }

        var minRow = cells.Min(c => c.Row);
        var maxRow = cells.Max(c => c.Row);
        var minCol = cells.Min(c => c.Column);
        var maxCol = cells.Max(c => c.Column);

        return JsonSerializer.Serialize(new
        {
            startCell = BuildCellReference(minCol, minRow),
            endCell = BuildCellReference(maxCol, maxRow),
            rowCount = (int)(maxRow - minRow + 1),
            columnCount = (int)(maxCol - minCol + 1)
        });
    }

    public void ExcelSetRangeValues(string sessionId, string sheetName, string startCell, string valuesJson)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(startCell);
        ArgumentException.ThrowIfNullOrWhiteSpace(valuesJson);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Excel);

        ExecuteWriteOperation(session, "excel_set_range_values", () =>
        {
            var matrixNode = JsonNode.Parse(valuesJson) ?? throw new InvalidOperationException("Invalid valuesJson.");
            var rows = matrixNode.AsArray();
            if (rows.Count == 0)
            {
                throw new InvalidOperationException("valuesJson must contain at least one row.");
            }

            var start = ParseCellReference(startCell);

            using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, true);
            var workbookPart = GetWorkbookPart(spreadsheet);
            var sheet = GetSheetByName(workbookPart, sheetName);
            var worksheetPart = GetWorksheetPart(workbookPart, sheet);
            var worksheet = GetWorksheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            for (var r = 0; r < rows.Count; r++)
            {
                var rowValues = rows[r]?.AsArray() ?? throw new InvalidOperationException("Each row must be an array.");
                var rowIndex = start.Row + (uint)r;
                var row = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex?.Value == rowIndex);
                if (row is null)
                {
                    row = new Row { RowIndex = rowIndex };
                    sheetData.Append(row);
                }

                for (var c = 0; c < rowValues.Count; c++)
                {
                    var value = rowValues[c]?.ToString() ?? string.Empty;
                    var columnIndex = start.Column + c;
                    var cellRef = BuildCellReference(columnIndex, rowIndex);

                    var cell = row.Elements<Cell>().FirstOrDefault(x => string.Equals(x.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                    if (cell is null)
                    {
                        cell = new Cell { CellReference = cellRef };
                        row.Append(cell);
                    }

                    cell.CellFormula = null;
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(value);
                }
            }

            worksheet.Save();
            GetWorkbook(workbookPart).Save();
        });
    }

    public void ExcelSetFormula(string sessionId, string sheetName, string cellReference, string formula)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        ArgumentException.ThrowIfNullOrWhiteSpace(formula);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Excel);

        ExecuteWriteOperation(session, "excel_set_formula", () =>
        {
            using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, true);
            var workbookPart = GetWorkbookPart(spreadsheet);
            var sheet = GetSheetByName(workbookPart, sheetName);
            var worksheetPart = GetWorksheetPart(workbookPart, sheet);
            var worksheet = GetWorksheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            var rowIndex = ParseRowIndex(cellReference);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
            if (row is null)
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            var normalizedRef = cellReference.ToUpperInvariant();
            var cell = row.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, normalizedRef, StringComparison.OrdinalIgnoreCase));
            if (cell is null)
            {
                cell = new Cell { CellReference = normalizedRef };
                row.Append(cell);
            }

            var normalizedFormula = formula.StartsWith('=') ? formula[1..] : formula;
            cell.CellFormula = new CellFormula(normalizedFormula);
            cell.CellValue = null;
            cell.DataType = null;

            worksheet.Save();
            GetWorkbook(workbookPart).Save();
        });
    }

    public string ExcelGetFormula(string sessionId, string sheetName, string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Excel)
        {
            throw new InvalidOperationException("Session document type is not Excel.");
        }

        using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheet = GetSheetByName(workbookPart, sheetName);
        var worksheet = GetWorksheet(GetWorksheetPart(workbookPart, sheet));
        var normalizedRef = cellReference.ToUpperInvariant();
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, normalizedRef, StringComparison.OrdinalIgnoreCase));

        return cell?.CellFormula?.Text ?? string.Empty;
    }

    public void PowerPointAddSlide(string sessionId, string title, string body)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(body);
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);

        ExecuteWriteOperation(session, "powerpoint_add_slide", () => PowerPointAddSlideCore(session.FilePath, title, body));
    }

    public void PowerPointInsertSlideAt(string sessionId, int index, string title, string body)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(body);
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);

        ExecuteWriteOperation(session, "powerpoint_insert_slide_at", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var presentationPart = GetPresentationPart(presentation);
            var slideIdList = GetOrCreateSlideIdList(presentationPart);
            var slides = slideIdList.Elements<SlideId>().ToList();

            if (index < 1 || index > slides.Count + 1)
            {
                throw new InvalidOperationException($"Index {index} is out of range. Valid range is 1..{slides.Count + 1}.");
            }

            var newSlidePart = presentationPart.AddNewPart<SlidePart>();
            newSlidePart.Slide = BuildSimpleSlide(title, body);
            var relationId = presentationPart.GetIdOfPart(newSlidePart);
            var newSlideIdValue = GetNextSlideId(slideIdList);
            var newSlideId = new SlideId { Id = newSlideIdValue, RelationshipId = relationId };

            if (index == slides.Count + 1)
            {
                slideIdList.Append(newSlideId);
            }
            else
            {
                slides[index - 1].InsertBeforeSelf(newSlideId);
            }

            newSlidePart.Slide.Save();
            presentationPart.Presentation?.Save();
        });
    }

    public void PowerPointSetSlideTitle(string sessionId, int slideIndex, string title)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        ExecuteWriteOperation(session, "powerpoint_set_slide_title", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
            SetSlideTextBySlot(slidePart, 0, title);
            GetSlide(slidePart).Save();
        });
    }

    public void PowerPointSetSlideBody(string sessionId, int slideIndex, string body)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(body);
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        ExecuteWriteOperation(session, "powerpoint_set_slide_body", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
            SetSlideTextBySlot(slidePart, 1, body);
            GetSlide(slidePart).Save();
        });
    }

    public void PowerPointReorderSlide(string sessionId, int fromIndex, int toIndex)
    {
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        ExecuteWriteOperation(session, "powerpoint_reorder_slide", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slideIdList = GetOrCreateSlideIdList(GetPresentationPart(presentation));
            var slides = slideIdList.Elements<SlideId>().ToList();

            if (fromIndex < 1 || fromIndex > slides.Count || toIndex < 1 || toIndex > slides.Count)
            {
                throw new InvalidOperationException($"Slide indices must be between 1 and {slides.Count}.");
            }

            if (fromIndex == toIndex)
            {
                return;
            }

            var moved = slides[fromIndex - 1];
            moved.Remove();
            slides.RemoveAt(fromIndex - 1);

            if (toIndex - 1 >= slides.Count)
            {
                slideIdList.Append(moved);
            }
            else
            {
                slides[toIndex - 1].InsertBeforeSelf(moved);
            }

            GetPresentationPart(presentation).Presentation?.Save();
        });
    }

    public void PowerPointDeleteSlide(string sessionId, int slideIndex)
    {
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        ExecuteWriteOperation(session, "powerpoint_delete_slide", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var presentationPart = GetPresentationPart(presentation);
            var slideIdList = GetOrCreateSlideIdList(presentationPart);
            var slides = slideIdList.Elements<SlideId>().ToList();

            if (slideIndex < 1 || slideIndex > slides.Count)
            {
                throw new InvalidOperationException($"Slide index {slideIndex} is out of range. Valid range is 1..{slides.Count}.");
            }

            var target = slides[slideIndex - 1];
            var slidePart = GetSlidePart(presentationPart, target);
            target.Remove();
            presentationPart.DeletePart(slidePart);
            presentationPart.Presentation?.Save();
        });
    }

    private static void PowerPointAddSlideCore(string filePath, string title, string body)
    {
        using var presentation = PresentationDocument.Open(filePath, true);
        var presentationPart = presentation.PresentationPart ?? throw new InvalidOperationException("Presentation part is missing.");
        var slideLayoutPart = EnsurePresentationDefaults(presentationPart);
        var presentationDocument = presentationPart.Presentation ?? throw new InvalidOperationException("Presentation document is missing.");
        var slideIdList = presentationDocument.SlideIdList ?? presentationDocument.AppendChild(new SlideIdList());

        var newSlidePart = presentationPart.AddNewPart<SlidePart>();
        _ = newSlidePart.AddPart(slideLayoutPart);
        newSlidePart.Slide = BuildSimpleSlide(title, body);

        var maxSlideId = slideIdList.Elements<SlideId>().Select(x => x.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max();
        var relationId = presentationPart.GetIdOfPart(newSlidePart);
        slideIdList.Append(new SlideId { Id = maxSlideId + 1, RelationshipId = relationId });

        newSlidePart.Slide.Save();
        presentationDocument.Save();
    }

    private static SlideLayoutPart EnsurePresentationDefaults(PresentationPart presentationPart)
    {
        var presentation = presentationPart.Presentation ??= new Presentation();
        presentation.SlideIdList ??= new SlideIdList();
        presentation.SlideSize ??= new SlideSize { Cx = 10080625, Cy = 7559675 };
        presentation.NotesSize ??= new NotesSize { Cx = 7559675, Cy = 10691813 };

        if (presentationPart.PresentationPropertiesPart is null)
        {
            var presPropsPart = presentationPart.AddNewPart<PresentationPropertiesPart>();
            presPropsPart.PresentationProperties = new PresentationProperties();
            presPropsPart.PresentationProperties.Save();
        }

        if (presentationPart.SlideMasterParts.Any())
        {
            return presentationPart.SlideMasterParts.First().SlideLayoutParts.First();
        }

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        var themePart = slideMasterPart.AddNewPart<ThemePart>();

        slideLayoutPart.SlideLayout = BuildDefaultSlideLayout();

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new SlideLayoutIdList(new SlideLayoutId
            {
                Id = 256U,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            }),
            new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));

        WriteDefaultTheme(themePart, ResolveThemeOptions("default"));
        _ = presentationPart.AddPart(themePart);

        presentation.SlideMasterIdList ??= new SlideMasterIdList();
        presentation.SlideMasterIdList.Append(new SlideMasterId
        {
            Id = 2147483648U,
            RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
        });

        slideLayoutPart.SlideLayout.Save();
        slideMasterPart.SlideMaster.Save();
        presentation.Save();

        return slideLayoutPart;
    }

    private static SlideLayout BuildDefaultSlideLayout()
    {
        var shapeTree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L },
                new A.ChildOffset { X = 0L, Y = 0L },
                new A.ChildExtents { Cx = 0L, Cy = 0L })));

        shapeTree.Append(CreateLayoutPlaceholderShape(
            id: 2U,
            name: "Title Placeholder",
            placeholderType: PlaceholderValues.Title,
            x: 685800L,
            y: 304800L,
            cx: 9144000L,
            cy: 1143000L));

        shapeTree.Append(CreateLayoutPlaceholderShape(
            id: 3U,
            name: "Content Placeholder",
            placeholderType: PlaceholderValues.Body,
            x: 685800L,
            y: 1828800L,
            cx: 9144000L,
            cy: 4114800L));

        return new SlideLayout(
            new CommonSlideData(shapeTree),
            new ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = SlideLayoutValues.Text,
            Preserve = true
        };
    }

    private static P.Shape CreateLayoutPlaceholderShape(
        uint id,
        string name,
        PlaceholderValues placeholderType,
        long x,
        long y,
        long cx,
        long cy)
    {
        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = placeholderType })),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.EndParagraphRunProperties())));
    }

    private static void WriteDefaultTheme(ThemePart themePart, PresentationThemeOptions options)
    {
        var theme = new A.Theme { Name = options.Name };
        var colorScheme = new A.ColorScheme { Name = "Default" };
        colorScheme.Append(
            new A.Dark1Color(new A.RgbColorModelHex { Val = options.Dark1 }),
            new A.Light1Color(new A.RgbColorModelHex { Val = options.Light1 }),
            new A.Dark2Color(new A.RgbColorModelHex { Val = options.Dark2 }),
            new A.Light2Color(new A.RgbColorModelHex { Val = options.Light2 }),
            new A.Accent1Color(new A.RgbColorModelHex { Val = options.Accent1 }),
            new A.Accent2Color(new A.RgbColorModelHex { Val = options.Accent2 }),
            new A.Accent3Color(new A.RgbColorModelHex { Val = options.Accent3 }),
            new A.Accent4Color(new A.RgbColorModelHex { Val = options.Accent4 }),
            new A.Accent5Color(new A.RgbColorModelHex { Val = options.Accent5 }),
            new A.Accent6Color(new A.RgbColorModelHex { Val = options.Accent6 }),
            new A.Hyperlink(new A.RgbColorModelHex { Val = options.Hyperlink }),
            new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = options.FollowedHyperlink }));

        var fontScheme = new A.FontScheme { Name = options.Name };
        fontScheme.Append(
            new A.MajorFont(new A.LatinFont { Typeface = options.MajorLatinFont }),
            new A.MinorFont(new A.LatinFont { Typeface = options.MinorLatinFont }));

        var formatScheme = new A.FormatScheme { Name = "Office" };
        formatScheme.Append(
            new A.FillStyleList(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
            new A.LineStyleList(
                new A.Outline(new A.PresetDash { Val = A.PresetLineDashValues.Solid }) { Width = 6350 },
                new A.Outline(new A.PresetDash { Val = A.PresetLineDashValues.Solid }) { Width = 12700 },
                new A.Outline(new A.PresetDash { Val = A.PresetLineDashValues.Solid }) { Width = 19050 }),
            new A.EffectStyleList(
                new A.EffectStyle(new A.EffectList()),
                new A.EffectStyle(new A.EffectList()),
                new A.EffectStyle(new A.EffectList())),
            new A.BackgroundFillStyleList(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })));

        theme.Append(new A.ThemeElements(colorScheme, fontScheme, formatScheme));
        themePart.Theme = theme;
        themePart.Theme.Save();
    }

    private static PresentationThemeOptions ResolveThemeOptions(string preset)
    {
        return preset switch
        {
            "neutral" => new PresentationThemeOptions(
                Name: "Neutral",
                Dark1: "1C1C1C",
                Light1: "FFFFFF",
                Dark2: "334155",
                Light2: "E2E8F0",
                Accent1: "2563EB",
                Accent2: "0EA5E9",
                Accent3: "10B981",
                Accent4: "F59E0B",
                Accent5: "EF4444",
                Accent6: "8B5CF6",
                Hyperlink: "1D4ED8",
                FollowedHyperlink: "6D28D9",
                MajorLatinFont: "Calibri",
                MinorLatinFont: "Calibri"),
            _ => new PresentationThemeOptions(
                Name: "Office",
                Dark1: "000000",
                Light1: "FFFFFF",
                Dark2: "1F497D",
                Light2: "EEECE1",
                Accent1: "4F81BD",
                Accent2: "C0504D",
                Accent3: "9BBB59",
                Accent4: "8064A2",
                Accent5: "4BACC6",
                Accent6: "F79646",
                Hyperlink: "0000FF",
                FollowedHyperlink: "800080",
                MajorLatinFont: "Arial",
                MinorLatinFont: "Arial")
        };
    }

    private static Slide BuildSimpleSlide(string title, string body)
    {
        var shapeTree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        shapeTree.Append(CreateTextShape(2U, "Title", title, 457200L));
        shapeTree.Append(CreateTextShape(3U, "Body", body, 1828800L));

        return new Slide(new CommonSlideData(shapeTree), new ColorMapOverride(new A.MasterColorMapping()));
    }

    private static Shape CreateTextShape(uint id, string name, string text, long yOffset)
    {
        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 457200L, Y = yOffset },
                    new A.Extents { Cx = 8229600L, Cy = 900000L })),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(text)), new A.EndParagraphRunProperties())));
    }

    private static SlideIdList GetOrCreateSlideIdList(PresentationPart presentationPart)
    {
        var presentation = presentationPart.Presentation ?? throw new InvalidOperationException("Presentation document is missing.");
        return presentation.SlideIdList ?? presentation.AppendChild(new SlideIdList());
    }

    private static uint GetNextSlideId(SlideIdList slideIdList)
    {
        return slideIdList.Elements<SlideId>().Select(x => x.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max() + 1;
    }

    private static SlidePart GetSlidePartByIndex(PresentationPart presentationPart, int slideIndex)
    {
        var slideIds = GetSlideIds(presentationPart).ToList();
        if (slideIndex < 1 || slideIndex > slideIds.Count)
        {
            throw new InvalidOperationException($"Slide index {slideIndex} is out of range. Valid range is 1..{slideIds.Count}.");
        }

        return GetSlidePart(presentationPart, slideIds[slideIndex - 1]);
    }

    private static void SetSlideTextBySlot(SlidePart slidePart, int slot, string text)
    {
        var paragraphs = GetSlide(slidePart)
            .Descendants<P.Shape>()
            .Select(shape => shape.TextBody)
            .Where(textBody => textBody is not null)
            .SelectMany(textBody => textBody!.Elements<A.Paragraph>())
            .ToList();

        if (paragraphs.Count == 0)
        {
            throw new InvalidOperationException("Slide does not contain editable text placeholders.");
        }

        var targetIndex = Math.Clamp(slot, 0, paragraphs.Count - 1);
        var paragraph = paragraphs[targetIndex];
        paragraph.RemoveAllChildren<A.Run>();
        paragraph.RemoveAllChildren<A.Break>();
        paragraph.RemoveAllChildren<A.Field>();
        paragraph.Append(new A.Run(new A.Text(text)));
    }

    private static void ApplyPowerPointSlotStyle(SlidePart slidePart, int slot, TextStyle style)
    {
        var paragraphs = GetSlide(slidePart)
            .Descendants<P.Shape>()
            .Select(shape => shape.TextBody)
            .Where(textBody => textBody is not null)
            .SelectMany(textBody => textBody!.Elements<A.Paragraph>())
            .ToList();

        if (paragraphs.Count == 0)
        {
            throw new InvalidOperationException("Slide does not contain editable text placeholders.");
        }

        var targetIndex = Math.Clamp(slot, 0, paragraphs.Count - 1);
        var paragraph = paragraphs[targetIndex];
        foreach (var run in paragraph.Elements<A.Run>())
        {
            run.RunProperties ??= new A.RunProperties();
            ApplyDrawingRunProperties(run.RunProperties, style);
        }

        var endParagraphProperties = paragraph.GetFirstChild<A.EndParagraphRunProperties>() ?? paragraph.AppendChild(new A.EndParagraphRunProperties());
        ApplyDrawingRunProperties(endParagraphProperties, style);
    }

    private static void ApplyWordParagraphStyle(W.Paragraph paragraph, TextStyle style)
    {
        foreach (var run in paragraph.Elements<W.Run>())
        {
            run.RunProperties ??= new W.RunProperties();
            ApplyWordRunProperties(run.RunProperties, style);
        }

        var paragraphProperties = paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        paragraphProperties.ParagraphMarkRunProperties ??= new W.ParagraphMarkRunProperties();
        ApplyWordParagraphMarkProperties(paragraphProperties.ParagraphMarkRunProperties, style);
    }

    private static void ApplyWordParagraphSpacing(W.Paragraph paragraph, int beforePt, int afterPt, double lineSpacing)
    {
        var paragraphProperties = paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        var spacing = paragraphProperties.SpacingBetweenLines ??= new W.SpacingBetweenLines();
        spacing.Before = (beforePt * 20).ToString();
        spacing.After = (afterPt * 20).ToString();
        spacing.Line = ((int)Math.Round(lineSpacing * 240)).ToString();
        spacing.LineRule = W.LineSpacingRuleValues.Auto;
    }

    private static void ApplyWordRunProperties(W.RunProperties runProperties, TextStyle style)
    {
        runProperties.RunFonts = new W.RunFonts { Ascii = style.FontName, HighAnsi = style.FontName, ComplexScript = style.FontName };
        runProperties.FontSize = new W.FontSize { Val = (style.FontSize * 2).ToString() };
        runProperties.Bold = style.Bold ? new W.Bold() : null;
        runProperties.Italic = style.Italic ? new W.Italic() : null;
        runProperties.Color = new W.Color { Val = style.ColorHex };
    }

    private static void ApplyWordParagraphMarkProperties(W.ParagraphMarkRunProperties markProperties, TextStyle style)
    {
        markProperties.RemoveAllChildren();
        markProperties.Append(new W.RunFonts { Ascii = style.FontName, HighAnsi = style.FontName, ComplexScript = style.FontName });
        markProperties.Append(new W.FontSize { Val = (style.FontSize * 2).ToString() });
        if (style.Bold)
        {
            markProperties.Append(new W.Bold());
        }

        if (style.Italic)
        {
            markProperties.Append(new W.Italic());
        }

        markProperties.Append(new W.Color { Val = style.ColorHex });
    }

    private static void ApplyDrawingRunProperties(A.TextCharacterPropertiesType runProperties, TextStyle style)
    {
        runProperties.FontSize = style.FontSize * 100;
        runProperties.Bold = style.Bold;
        runProperties.Italic = style.Italic;
        runProperties.RemoveAllChildren<A.LatinFont>();
        runProperties.Append(new A.LatinFont { Typeface = style.FontName });
        runProperties.RemoveAllChildren<A.SolidFill>();
        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = style.ColorHex }));
    }

    private static TextStyle BuildTextStyleFromPreset(string preset, bool isTitleSlot)
    {
        return preset switch
        {
            "title" => isTitleSlot
                ? new TextStyle("Calibri", 36, true, false, "1F497D")
                : new TextStyle("Calibri", 22, false, false, "4F4F4F"),
            "subtitle" => isTitleSlot
                ? new TextStyle("Calibri", 28, true, false, "1F497D")
                : new TextStyle("Calibri", 20, false, true, "666666"),
            _ => new TextStyle("Calibri", 18, false, false, "000000")
        };
    }

    private static void ValidateTextStyle(TextStyle style)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(style.FontName);
        if (style.FontSize < 8 || style.FontSize > 96)
        {
            throw new InvalidOperationException("Font size must be between 8 and 96.");
        }

        if (!IsValidHexColor(style.ColorHex))
        {
            throw new InvalidOperationException("Color must be a 6-digit hex value like FF0000.");
        }
    }

    private static bool IsValidHexColor(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length != 6)
        {
            return false;
        }

        return value.All(Uri.IsHexDigit);
    }

    private static string SanitizeStyleId(string styleName)
    {
        var cleaned = new string(styleName.Where(char.IsLetterOrDigit).ToArray());
        return string.IsNullOrWhiteSpace(cleaned) ? $"Custom{Guid.NewGuid():N}" : cleaned;
    }

    private static void ApplyStyleOptions(W.Style style, JsonObject options)
    {
        if (!style.Elements<W.UnhideWhenUsed>().Any())
        {
            style.Append(new W.UnhideWhenUsed());
        }

        if (!style.Elements<W.PrimaryStyle>().Any())
        {
            style.Append(new W.PrimaryStyle());
        }

        if (!style.Elements<W.UIPriority>().Any())
        {
            style.Append(new W.UIPriority { Val = 99 });
        }

        var paragraphProperties = style.StyleParagraphProperties ??= new W.StyleParagraphProperties();
        var runProperties = style.StyleRunProperties ??= new W.StyleRunProperties();

        if (options.TryGetPropertyValue("basedOn", out var basedOnNode) && basedOnNode is not null)
        {
            style.BasedOn = new W.BasedOn { Val = basedOnNode.GetValue<string>() };
        }

        if (options.TryGetPropertyValue("nextStyle", out var nextStyleNode) && nextStyleNode is not null)
        {
            style.NextParagraphStyle = new W.NextParagraphStyle { Val = nextStyleNode.GetValue<string>() };
        }

        if (options.TryGetPropertyValue("fontName", out var fontNameNode) && fontNameNode is not null)
        {
            var fontName = fontNameNode.GetValue<string>();
            runProperties.RunFonts = new W.RunFonts { Ascii = fontName, HighAnsi = fontName, ComplexScript = fontName };
        }

        if (options.TryGetPropertyValue("fontSize", out var fontSizeNode) && fontSizeNode is not null)
        {
            runProperties.FontSize = new W.FontSize { Val = (fontSizeNode.GetValue<int>() * 2).ToString() };
        }

        if (options.TryGetPropertyValue("bold", out var boldNode) && boldNode is not null)
        {
            runProperties.Bold = boldNode.GetValue<bool>() ? new W.Bold() : null;
        }

        if (options.TryGetPropertyValue("italic", out var italicNode) && italicNode is not null)
        {
            runProperties.Italic = italicNode.GetValue<bool>() ? new W.Italic() : null;
        }

        if (options.TryGetPropertyValue("colorHex", out var colorNode) && colorNode is not null)
        {
            var color = colorNode.GetValue<string>();
            if (!IsValidHexColor(color))
            {
                throw new InvalidOperationException("Style colorHex must be a 6-digit hex value.");
            }

            runProperties.Color = new W.Color { Val = color };
        }

        var beforePt = options.TryGetPropertyValue("beforePt", out var beforeNode) && beforeNode is not null ? beforeNode.GetValue<int>() : 0;
        var afterPt = options.TryGetPropertyValue("afterPt", out var afterNode) && afterNode is not null ? afterNode.GetValue<int>() : 0;
        var lineSpacing = options.TryGetPropertyValue("lineSpacing", out var lineSpacingNode) && lineSpacingNode is not null ? lineSpacingNode.GetValue<double>() : 0;
        if (beforePt > 0 || afterPt > 0 || lineSpacing > 0)
        {
            paragraphProperties.SpacingBetweenLines ??= new W.SpacingBetweenLines();
            if (beforePt >= 0)
            {
                paragraphProperties.SpacingBetweenLines.Before = (beforePt * 20).ToString();
            }

            if (afterPt >= 0)
            {
                paragraphProperties.SpacingBetweenLines.After = (afterPt * 20).ToString();
            }

            if (lineSpacing > 0)
            {
                paragraphProperties.SpacingBetweenLines.Line = ((int)Math.Round(lineSpacing * 240)).ToString();
                paragraphProperties.SpacingBetweenLines.LineRule = W.LineSpacingRuleValues.Auto;
            }
        }
    }

    private static uint EnsureExcelCellFormat(WorkbookPart workbookPart, TextStyle style)
    {
        var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet ??= new Stylesheet(
            new DocumentFormat.OpenXml.Spreadsheet.Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()) { Count = 1U },
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2U },
            new Borders(new DocumentFormat.OpenXml.Spreadsheet.Border()) { Count = 1U },
            new CellStyleFormats(new CellFormat()) { Count = 1U },
            new CellFormats(new CellFormat()) { Count = 1U });

        var stylesheet = stylesPart.Stylesheet;
        stylesheet.Fonts ??= new DocumentFormat.OpenXml.Spreadsheet.Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()) { Count = 1U };
        stylesheet.CellFormats ??= new CellFormats(new CellFormat()) { Count = 1U };

        var font = new DocumentFormat.OpenXml.Spreadsheet.Font(
            new FontName { Val = style.FontName },
            new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = style.FontSize },
            new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = style.ColorHex });
        if (style.Bold)
        {
            font.Append(new DocumentFormat.OpenXml.Spreadsheet.Bold());
        }

        if (style.Italic)
        {
            font.Append(new DocumentFormat.OpenXml.Spreadsheet.Italic());
        }

        stylesheet.Fonts.Append(font);
        stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

        var fontId = stylesheet.Fonts.Count.Value - 1;
        var format = new CellFormat { FontId = fontId, ApplyFont = true };
        stylesheet.CellFormats.Append(format);
        stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

        stylesheet.Save();
        return stylesheet.CellFormats.Count.Value - 1;
    }

    private static object ListWordStructure(string filePath)
    {
        using var document = WordprocessingDocument.Open(filePath, false);
        var body = document.MainDocumentPart?.Document?.Body;
        var paragraphCount = body?.Descendants<Paragraph>().Count() ?? 0;
        var tableCount = body?.Descendants<W.Table>().Count() ?? 0;
        return new { paragraphCount, tableCount };
    }

    private static object ListExcelStructure(string filePath)
    {
        using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheetNames = GetSheets(workbookPart).Select(s => s.Name?.Value ?? string.Empty).ToArray();
        return new { sheetCount = sheetNames.Length, sheetNames };
    }

    private static object ListPowerPointStructure(string filePath)
    {
        using var presentation = PresentationDocument.Open(filePath, false);
        var slideCount = GetSlideIds(GetPresentationPart(presentation)).Count();
        return new { slideCount };
    }

    private static object FindTextInWord(string filePath, string query)
    {
        using var document = WordprocessingDocument.Open(filePath, false);
        var paragraphs = document.MainDocumentPart?.Document?.Body?.Descendants<Paragraph>() ?? [];
        var matches = paragraphs
            .Select((p, index) => new { index = index + 1, text = p.InnerText })
            .Where(x => x.text.Contains(query, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        return new { matchCount = matches.Length, matches };
    }

    private static object FindTextInExcel(string filePath, string query)
    {
        using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        var matches = new List<object>();

        foreach (var sheet in GetSheets(workbookPart))
        {
            var worksheetPart = GetWorksheetPart(workbookPart, sheet);
            foreach (var cell in GetWorksheet(worksheetPart).Descendants<Cell>())
            {
                var value = ReadCellText(cell, sharedStrings);
                if (!string.IsNullOrEmpty(value) && value.Contains(query, StringComparison.OrdinalIgnoreCase))
                {
                    matches.Add(new { sheet = sheet.Name?.Value ?? string.Empty, cell = cell.CellReference?.Value ?? string.Empty, value });
                }
            }
        }

        return new { matchCount = matches.Count, matches };
    }

    private static object FindTextInPowerPoint(string filePath, string query)
    {
        using var presentation = PresentationDocument.Open(filePath, false);
        var presentationPart = GetPresentationPart(presentation);
        var slideIds = GetSlideIds(presentationPart);
        var matches = new List<object>();

        var index = 0;
        foreach (var slideId in slideIds)
        {
            index++;
            var slidePart = GetSlidePart(presentationPart, slideId);
            var textContent = string.Join(" ", GetSlide(slidePart).Descendants<A.Text>().Select(t => t.Text));
            if (textContent.Contains(query, StringComparison.OrdinalIgnoreCase))
            {
                matches.Add(new { slideIndex = index, text = textContent });
            }
        }

        return new { matchCount = matches.Count, matches };
    }

    private static string ReadCellText(Cell cell, SharedStringTable? sharedStrings)
    {
        var rawValue = cell.CellValue?.Text ?? string.Empty;
        if (cell.DataType?.Value == CellValues.SharedString && int.TryParse(rawValue, out var sharedIndex))
        {
            return sharedStrings?.ElementAtOrDefault(sharedIndex)?.InnerText ?? string.Empty;
        }

        return rawValue;
    }

    private static int CountOccurrences(string input, string value, StringComparison comparison)
    {
        if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(value))
        {
            return 0;
        }

        var count = 0;
        var startIndex = 0;
        while (true)
        {
            var index = input.IndexOf(value, startIndex, comparison);
            if (index < 0)
            {
                break;
            }

            count++;
            startIndex = index + Math.Max(1, value.Length);
        }

        return count;
    }

    private static string ReplaceWithComparison(string input, string oldValue, string newValue, StringComparison comparison)
    {
        if (comparison == StringComparison.Ordinal)
        {
            return input.Replace(oldValue, newValue, StringComparison.Ordinal);
        }

        var startIndex = 0;
        var result = new System.Text.StringBuilder();
        while (true)
        {
            var index = input.IndexOf(oldValue, startIndex, comparison);
            if (index < 0)
            {
                result.Append(input.AsSpan(startIndex));
                break;
            }

            result.Append(input.AsSpan(startIndex, index - startIndex));
            result.Append(newValue);
            startIndex = index + oldValue.Length;
        }

        return result.ToString();
    }

    private static ListNumberingDefinitions EnsureStructuredListNumbering(WordprocessingDocument document, IReadOnlyCollection<WordListItem> items)
    {
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
        var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new W.Numbering();
        var numbering = numberingPart.Numbering;

        const int numberedAbstractId = 901;
        const int bulletedAbstractId = 902;
        const int numberedNumberingId = 901;
        const int bulletedNumberingId = 902;

        var maxLevel = items.Max(x => x.Level);
        EnsureAbstractNumbering(numbering, numberedAbstractId, maxLevel, isNumbered: true, items);
        EnsureAbstractNumbering(numbering, bulletedAbstractId, maxLevel, isNumbered: false, items);

        EnsureNumberingInstance(numbering, numberedNumberingId, numberedAbstractId);
        EnsureNumberingInstance(numbering, bulletedNumberingId, bulletedAbstractId);

        numbering.Save();
        return new ListNumberingDefinitions(numberedNumberingId, bulletedNumberingId);
    }

    private static void EnsureAbstractNumbering(W.Numbering numbering, int abstractId, int maxLevel, bool isNumbered, IReadOnlyCollection<WordListItem> items)
    {
        var existing = numbering.Elements<W.AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == abstractId);
        existing?.Remove();

        var abstractNum = new W.AbstractNum { AbstractNumberId = abstractId };
        for (var level = 0; level <= maxLevel; level++)
        {
            var style = isNumbered
                ? ResolveNumberStyle(items, level)
                : ResolveBulletStyle(items, level);
            abstractNum.Append(BuildListLevel(level, style, isNumbered));
        }

        numbering.Append(abstractNum);
    }

    private static void EnsureNumberingInstance(W.Numbering numbering, int numberingId, int abstractId)
    {
        var existing = numbering.Elements<W.NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numberingId);
        existing?.Remove();
        numbering.Append(new W.NumberingInstance(new W.AbstractNumId { Val = abstractId }) { NumberID = numberingId });
    }

    private static W.Level BuildListLevel(int level, string style, bool isNumbered)
    {
        var left = 720 * (level + 1);
        var hanging = 360;
        var numberingFormat = isNumbered ? ResolveNumberFormat(style) : W.NumberFormatValues.Bullet;
        var levelText = isNumbered ? ResolveNumberLevelText(style, level) : ResolveBulletGlyph(style);

        return new W.Level(
            new W.StartNumberingValue { Val = 1 },
            new W.NumberingFormat { Val = numberingFormat },
            new W.LevelText { Val = levelText },
            new W.LevelJustification { Val = W.LevelJustificationValues.Left },
            new W.ParagraphProperties(new W.Indentation { Left = left.ToString(), Hanging = hanging.ToString() }),
            new W.RunProperties(new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }))
        { LevelIndex = level };
    }

    private static string ResolveNumberStyle(IReadOnlyCollection<WordListItem> items, int level)
    {
        return items.FirstOrDefault(x => x.Level == level && NormalizeListKind(x.Kind) == "numbered")?.NumberStyle
            ?.Trim().ToLowerInvariant()
            ?? "decimal-dot";
    }

    private static string ResolveBulletStyle(IReadOnlyCollection<WordListItem> items, int level)
    {
        return items.FirstOrDefault(x => x.Level == level && NormalizeListKind(x.Kind) == "bulleted")?.BulletStyle
            ?.Trim().ToLowerInvariant()
            ?? "disc";
    }

    private static string NormalizeListKind(string kind)
    {
        var normalized = kind?.Trim().ToLowerInvariant();
        return normalized is "numbered" or "ordered" ? "numbered" : "bulleted";
    }

    private static W.NumberFormatValues ResolveNumberFormat(string style)
    {
        return style switch
        {
            "upper-alpha-dot" => W.NumberFormatValues.UpperLetter,
            "lower-alpha-paren" => W.NumberFormatValues.LowerLetter,
            "upper-roman-dot" => W.NumberFormatValues.UpperRoman,
            "lower-roman-dot" => W.NumberFormatValues.LowerRoman,
            _ => W.NumberFormatValues.Decimal
        };
    }

    private static string ResolveNumberLevelText(string style, int level)
    {
        var marker = $"%{level + 1}";
        return style switch
        {
            "lower-alpha-paren" => $"{marker})",
            "decimal-paren" => $"{marker})",
            _ => $"{marker}."
        };
    }

    private static string ResolveBulletGlyph(string style)
    {
        return style switch
        {
            "circle" => "o",
            "square" => "■",
            "diamond" => "◆",
            "dash" => "-",
            "arrow" => "➤",
            "check" => "✓",
            _ => "•"
        };
    }

    private void ExecuteBatchOperation(string sessionId, string operation, JsonNode payload)
    {
        OfficeBatchDispatcher.Dispatch(this, sessionId, operation, payload);
    }

    private static string NormalizeOperationName(string operationName)
    {
        return operationName.Trim().ToLowerInvariant();
    }

    private static OfficeOperationSpec? TryGetOperationSpec(string normalizedOperation)
    {
        return OfficeOperationRegistry.TryGet(normalizedOperation);
    }

    private static string GetErrorCode(Exception ex)
    {
        return ex switch
        {
            InvalidOperationException => "InvalidOperation",
            ArgumentException => "InvalidArgument",
            JsonException => "InvalidJson",
            _ => "UnhandledError"
        };
    }

    private OfficeSession GetSession(string sessionId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sessionId);
        if (_sessions.TryGetValue(sessionId, out var session))
        {
            return session;
        }

        throw new InvalidOperationException($"Session '{sessionId}' was not found.");
    }

    private static WorkbookPart GetWorkbookPart(SpreadsheetDocument spreadsheet)
    {
        return spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
    }

    private static IEnumerable<Sheet> GetSheets(WorkbookPart workbookPart)
    {
        var sheets = GetWorkbook(workbookPart).Sheets;
        if (sheets is null)
        {
            throw new InvalidOperationException("Workbook sheets are missing.");
        }

        return sheets.Elements<Sheet>();
    }

    private static Sheet GetSheetByName(WorkbookPart workbookPart, string sheetName)
    {
        return GetSheets(workbookPart).FirstOrDefault(s => s.Name?.Value == sheetName)
            ?? throw new InvalidOperationException($"Sheet '{sheetName}' was not found.");
    }

    private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, Sheet sheet)
    {
        var relationshipId = sheet.Id?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            throw new InvalidOperationException("Sheet relationship id is missing.");
        }

        return (WorksheetPart)workbookPart.GetPartById(relationshipId);
    }

    private static Workbook GetWorkbook(WorkbookPart workbookPart)
    {
        return workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
    }

    private static Worksheet GetWorksheet(WorksheetPart worksheetPart)
    {
        return worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
    }

    private static PresentationPart GetPresentationPart(PresentationDocument presentation)
    {
        return presentation.PresentationPart ?? throw new InvalidOperationException("Presentation part is missing.");
    }

    private static IEnumerable<SlideId> GetSlideIds(PresentationPart presentationPart)
    {
        return presentationPart.Presentation?.SlideIdList?.Elements<SlideId>() ?? [];
    }

    private static SlidePart GetSlidePart(PresentationPart presentationPart, SlideId slideId)
    {
        var relationshipId = slideId.RelationshipId?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            throw new InvalidOperationException("Slide relationship id is missing.");
        }

        return (SlidePart)presentationPart.GetPartById(relationshipId);
    }

    private static Slide GetSlide(SlidePart slidePart)
    {
        return slidePart.Slide ?? throw new InvalidOperationException("Slide content is missing.");
    }

    private OfficeSession GetWritableSession(string sessionId, OfficeDocumentType requiredType)
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

    private static void CreateEmptyDocument(string filePath, OfficeDocumentType documentType)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }

        switch (documentType)
        {
            case OfficeDocumentType.Word:
                using (var word = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                {
                    var mainPart = word.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());
                }
                break;
            case OfficeDocumentType.Excel:
                using (var excel = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = excel.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    sheets.Append(new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1U,
                        Name = "Sheet1"
                    });
                    workbookPart.Workbook.Save();
                }
                break;
            case OfficeDocumentType.PowerPoint:
                InitializePowerPointFromTemplate(filePath);
                break;
            default:
                throw new InvalidOperationException("Unsupported document type.");
        }
    }

    private static void InitializePowerPointFromTemplate(string filePath)
    {
        using (var presentation = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
        {
            var presentationPart = presentation.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            _ = EnsurePresentationDefaults(presentationPart);
            presentationPart.Presentation.Save();
        }
    }

    private static OfficeDocumentType ParseDocumentType(string documentType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(documentType);
        return documentType.Trim().ToLowerInvariant() switch
        {
            "docx" or "word" => OfficeDocumentType.Word,
            "xlsx" or "excel" => OfficeDocumentType.Excel,
            "pptx" or "powerpoint" or "ppt" => OfficeDocumentType.PowerPoint,
            _ => throw new InvalidOperationException($"Unsupported document type '{documentType}'.")
        };
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

    private static string NormalizePath(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        return Path.GetFullPath(filePath);
    }

    private void ExecuteWriteOperation(OfficeSession session, string operationName, Action operation)
    {
        var snapshotPath = CreateSnapshot(session);

        try
        {
            operation();
            AppendOperationLog(session.Id, operationName, "Operation completed.");
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

    private string CreateSnapshot(OfficeSession session)
    {
        var snapshotPath = Path.Combine(Path.GetTempPath(), $"openxmlmcp-snapshot-{session.Id}-{Guid.NewGuid():N}{Path.GetExtension(session.FilePath)}");
        File.Copy(session.FilePath, snapshotPath, overwrite: true);

        var snapshots = _sessionSnapshots.GetOrAdd(session.Id, _ => new Stack<string>());
        snapshots.Push(snapshotPath);
        AppendOperationLog(session.Id, "checkpoint", $"Snapshot created: {snapshotPath}");
        return snapshotPath;
    }

    private void AppendOperationLog(string sessionId, string operationName, string message)
    {
        var history = _sessionOperationLog.GetOrAdd(sessionId, _ => new List<OperationLogEntry>());
        history.Add(new OperationLogEntry(DateTimeOffset.UtcNow, operationName, message));
    }

    private static void DeleteIfExists(string filePath)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }

    private static uint ParseRowIndex(string cellReference)
    {
        var rowText = new string(cellReference.Where(char.IsDigit).ToArray());
        if (uint.TryParse(rowText, out var rowIndex) && rowIndex > 0)
        {
            return rowIndex;
        }

        throw new InvalidOperationException($"Invalid cell reference '{cellReference}'.");
    }

    private static (int Column, uint Row) ParseCellReference(string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);

        var letters = new string(cellReference.Where(char.IsLetter).ToArray()).ToUpperInvariant();
        var numbers = new string(cellReference.Where(char.IsDigit).ToArray());
        if (letters.Length == 0 || !uint.TryParse(numbers, out var row) || row == 0)
        {
            throw new InvalidOperationException($"Invalid cell reference '{cellReference}'.");
        }

        var column = 0;
        foreach (var ch in letters)
        {
            column = (column * 26) + (ch - 'A' + 1);
        }

        return (column, row);
    }

    private static string BuildCellReference(int columnIndex, uint rowIndex)
    {
        if (columnIndex <= 0 || rowIndex == 0)
        {
            throw new InvalidOperationException("Invalid row or column index.");
        }

        var column = new System.Text.StringBuilder();
        var current = columnIndex;
        while (current > 0)
        {
            current--;
            column.Insert(0, (char)('A' + (current % 26)));
            current /= 26;
        }

        return $"{column}{rowIndex}";
    }

    private sealed record ListNumberingDefinitions(int NumberedNumberingId, int BulletedNumberingId);
    private sealed record OperationLogEntry(DateTimeOffset TimestampUtc, string OperationName, string Message);
}
