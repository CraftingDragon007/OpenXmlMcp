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

    [Fact]
    public void WordSession_AddTable_IncreasesTableCount()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 2, 3);

            var structure = service.ListStructure(sessionId);
            Assert.Contains("\"tableCount\":1", structure, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ExcelSession_AddWorksheet_UpdatesStructure()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            service.ExcelAddWorksheet(sessionId, "Data");

            var structure = service.ListStructure(sessionId);
            Assert.Contains("\"sheetCount\":2", structure, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Data", structure);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ExcelSession_AddWorksheetThenSetValue_Works()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            service.ExcelAddWorksheet(sessionId, "Data");
            service.ExcelSetCellValue(sessionId, "Data", "C3", "PHASE2-OK");

            var value = service.ExcelGetCellValue(sessionId, "Data", "C3");
            Assert.Equal("PHASE2-OK", value);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void PowerPointSession_AddBulletSlide_IsSearchable()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddBulletSlide(sessionId, "Sprint", "Task A\nTask B");

            var result = service.FindText(sessionId, "Task B");
            Assert.Contains("\"matchCount\":1", result, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void BatchExecute_MixedOperations_ReportsFailuresAndSuccesses()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var operations = "[" +
                "{\"operation\":\"word_append_paragraph\",\"text\":\"hello\"}," +
                "{\"operation\":\"word_add_table\",\"rows\":1,\"columns\":1}," +
                "{\"operation\":\"excel_set_cell_value\",\"sheetName\":\"Sheet1\",\"cellReference\":\"A1\",\"value\":\"x\"}" +
                "]";

            var payload = service.BatchExecute(sessionId, operations);
            Assert.Contains("\"executed\":2", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"failed\":1", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ValidateOperation_KnowsPhase2OperationNames()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var payload = service.ValidateOperation(sessionId, "word_add_table");

            Assert.Contains("\"isValid\":true", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }
}
