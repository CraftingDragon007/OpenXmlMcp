using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlMcp.Server.Services;

namespace OpenXmlMcp.Server.Tests;

public partial class OfficeSessionServiceTests
{
    [Fact]
    public void ExcelSession_SetAndGetCell_Works()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

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
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ExcelSession_AddWorksheet_UpdatesStructure()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

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
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ExcelSession_AddWorksheetThenSetValue_Works()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

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
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ExcelSetRangeValues_AndGetUsedRange_Works()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            service.ExcelSetRangeValues(sessionId, "Sheet1", "B2", "[[\"A\",\"B\"],[\"C\",\"D\"]]");

            var usedRange = service.ExcelGetUsedRange(sessionId, "Sheet1");
            Assert.Contains("\"startCell\":\"B2\"", usedRange, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"endCell\":\"C3\"", usedRange, StringComparison.OrdinalIgnoreCase);

            var value = service.ExcelGetCellValue(sessionId, "Sheet1", "C3");
            Assert.Equal("D", value);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ExcelSetRangeValues_InvalidJson_ReturnsHelpfulError()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            var ex = Assert.Throws<InvalidOperationException>(() =>
                service.ExcelSetRangeValues(sessionId, "Sheet1", "A1", "[['A','B']]")
            );

            Assert.Contains("strict JSON", ex.Message, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ExcelSetRangeValues_StringFormula_IsStoredAsFormula()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            service.ExcelSetRangeValues(sessionId, "Sheet1", "A1", "[[\"=SUM(B1:C1)\"]]");

            var formula = service.ExcelGetFormula(sessionId, "Sheet1", "A1");
            Assert.Equal("SUM(B1:C1)", formula);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ExcelSetFormula_AndGetFormula_Works()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            service.ExcelSetFormula(sessionId, "Sheet1", "C1", "=A1+B1");

            var formula = service.ExcelGetFormula(sessionId, "Sheet1", "C1");
            Assert.Equal("A1+B1", formula);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ExcelSetCellStyle_AppliesStyleIndex()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            service.ExcelSetCellValue(sessionId, "Sheet1", "A1", "Styled");
            service.ExcelSetCellStyle(sessionId, "Sheet1", "A1", "Calibri", 12, bold: true, italic: false, colorHex: "112233");
            service.CloseDocument(sessionId);

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var worksheet = spreadsheet.WorkbookPart!
                .WorksheetParts
                .First()
                .Worksheet!;
            var cell = worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                .First(c => string.Equals(c.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));

            Assert.True(cell.StyleIndex?.Value > 0);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ExcelGetCellInfo_ReturnsFormulaAndCachedValueFields()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            service.ExcelSetFormula(sessionId, "Sheet1", "B2", "=A1+A2");

            var payload = service.ExcelGetCellInfo(sessionId, "Sheet1", "B2");
            using var doc = JsonDocument.Parse(payload);

            Assert.True(doc.RootElement.GetProperty("exists").GetBoolean());
            Assert.Equal("A1+A2", doc.RootElement.GetProperty("formula").GetString());
            Assert.True(doc.RootElement.TryGetProperty("cachedValue", out _));

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ExcelGetCellStyle_ReturnsResolvedFontFields()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            service.ExcelSetCellStyle(sessionId, "Sheet1", "A1", "Calibri", 12, true, false, "FF0000");

            var payload = service.ExcelGetCellStyle(sessionId, "Sheet1", "A1");
            Assert.Contains("\"exists\":true", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"fontName\"", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
}
