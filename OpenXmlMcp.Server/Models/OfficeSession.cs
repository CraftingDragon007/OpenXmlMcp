namespace OpenXmlMcp.Server.Models;

public sealed class OfficeSession
{
    public required string Id { get; init; }
    public required string FilePath { get; init; }
    public required OfficeDocumentType DocumentType { get; init; }
    public required bool IsReadOnly { get; init; }
    public required DateTimeOffset CreatedAtUtc { get; init; }
}
