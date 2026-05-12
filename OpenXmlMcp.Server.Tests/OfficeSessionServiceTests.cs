using System.Text.Json;
using OpenXmlMcp.Server.Services;

namespace OpenXmlMcp.Server.Tests;

public class OfficeSessionServiceTests
{
    [Fact]
    public void WordSession_CreateAppendFind_Works()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Hello MCP Word");

            var structure = service.ListStructure(sessionId);
            Assert.Contains("paragraphCount", structure);

            var findResult = service.FindText(sessionId, "mcp word");
            Assert.Contains("matchCount", findResult);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ExcelSession_SetAndGetCell_Works()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            service.ExcelSetCellValue(sessionId, "Sheet1", "A1", "42");

            var value = service.ExcelGetCellValue(sessionId, "Sheet1", "A1");
            Assert.Equal("42", value);

            var findResult = service.FindText(sessionId, "42");
            Assert.Contains("matchCount", findResult);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void PowerPointSession_AddSlideAndFind_Works()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Roadmap", "Phase one complete");

            var structure = service.ListStructure(sessionId);
            Assert.Contains("slideCount", structure);

            var findResult = service.FindText(sessionId, "phase one");
            Assert.Contains("matchCount", findResult);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void GetDocumentInfo_ReturnsSessionPayload()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var payload = service.GetDocumentInfo(sessionId);

            using var doc = JsonDocument.Parse(payload);
            Assert.Equal(sessionId, doc.RootElement.GetProperty("Id").GetString());
            Assert.Equal("Word", doc.RootElement.GetProperty("documentType").GetString());

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ValidateOperation_ReadOnlySession_ReturnsInvalidForWriteOperation()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var writableSessionId = service.CreateDocument(filePath, "docx");
            service.CloseDocument(writableSessionId);

            var readOnlySessionId = service.OpenDocument(filePath, readOnly: true);
            var payload = service.ValidateOperation(readOnlySessionId, "word_append_paragraph");

            Assert.Contains("\"isValid\":false", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("read-only", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(readOnlySessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    private static string GetTempPath(string extension)
    {
        return Path.Combine(Path.GetTempPath(), $"openxmlmcp-{Guid.NewGuid():N}.{extension}");
    }

    private static void DeleteIfExists(string filePath)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }
}
