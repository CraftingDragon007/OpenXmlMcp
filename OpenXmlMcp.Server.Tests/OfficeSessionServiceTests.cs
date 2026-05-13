using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
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

    [Fact]
    public void UndoLastChange_RestoresPreviousState()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Before undo");
            service.UndoLastChange(sessionId);

            var result = service.FindText(sessionId, "Before undo");
            Assert.Contains("\"matchCount\":0", result, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void GetOperationHistory_ReturnsEntries()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "History entry");

            var history = service.GetOperationHistory(sessionId);
            Assert.Contains("word_append_paragraph", history, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void OpenDocument_LargeFile_ThrowsSafetyException()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
            stream.SetLength(21L * 1024 * 1024);

            var ex = Assert.Throws<InvalidOperationException>(() => service.OpenDocument(filePath));
            Assert.Contains("exceeds safety limit", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordInsertParagraphAt_UsesOneBasedIndex()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "first");
            service.WordAppendParagraph(sessionId, "third");
            service.WordInsertParagraphAt(sessionId, 2, "second");

            var matches = service.FindText(sessionId, "second");
            Assert.Contains("\"index\":2", matches, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordReplaceText_ReturnsReplacementCount()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "alpha beta alpha");

            var count = service.WordReplaceText(sessionId, "alpha", "gamma", matchCase: false);

            Assert.Equal(2, count);
            var result = service.FindText(sessionId, "gamma");
            Assert.Contains("\"matchCount\":1", result, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAddHeading_AndBulletedList_AreSearchable()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddHeading(sessionId, 2, "Roadmap");
            service.WordAddBulletedList(sessionId, "Item A\nItem B");

            var headingResult = service.FindText(sessionId, "Roadmap");
            var bulletResult = service.FindText(sessionId, "Item B");
            Assert.Contains("\"matchCount\":1", headingResult, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"matchCount\":1", bulletResult, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ExcelSetRangeValues_AndGetUsedRange_Works()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("xlsx");

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
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ExcelSetFormula_AndGetFormula_Works()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("xlsx");

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
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void PowerPointInsertSlideAt_UsesOneBasedIndex()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "One", "First");
            service.PowerPointAddSlide(sessionId, "Three", "Third");
            service.PowerPointInsertSlideAt(sessionId, 2, "Two", "Second");

            var search = service.FindText(sessionId, "Two");
            Assert.Contains("\"slideIndex\":2", search, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void PowerPointSetSlideTitleAndBody_UpdatesSlideContent()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Old Title", "Old Body");
            service.PowerPointSetSlideTitle(sessionId, 1, "New Title");
            service.PowerPointSetSlideBody(sessionId, 1, "New Body");

            var titleSearch = service.FindText(sessionId, "New Title");
            var bodySearch = service.FindText(sessionId, "New Body");
            Assert.Contains("\"matchCount\":1", titleSearch, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"matchCount\":1", bodySearch, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void PowerPointReorderAndDeleteSlide_Work()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "One", "A");
            service.PowerPointAddSlide(sessionId, "Two", "B");
            service.PowerPointAddSlide(sessionId, "Three", "C");

            service.PowerPointReorderSlide(sessionId, 3, 1);
            var movedSearch = service.FindText(sessionId, "Three");
            Assert.Contains("\"slideIndex\":1", movedSearch, StringComparison.OrdinalIgnoreCase);

            service.PowerPointDeleteSlide(sessionId, 2);
            var structure = service.ListStructure(sessionId);
            Assert.Contains("\"slideCount\":2", structure, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void PowerPointCreateDocument_HasSlideMasterAndTheme()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.CloseDocument(sessionId);

            using var presentation = PresentationDocument.Open(filePath, false);
            var presentationPart = presentation.PresentationPart;

            Assert.NotNull(presentationPart);
            Assert.NotEmpty(presentationPart!.SlideMasterParts);
            var slideMasterPart = presentationPart.SlideMasterParts.First();
            Assert.NotNull(slideMasterPart.ThemePart);
            Assert.NotEmpty(slideMasterPart.SlideLayoutParts);
            Assert.Equal("Office", slideMasterPart.ThemePart!.Theme?.Name?.Value);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ValidateOperation_AcceptsPowerPointUnderscoreAliases()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");

            var validateReorder = service.ValidateOperation(sessionId, "power_point_reorder_slide");
            var validateDelete = service.ValidateOperation(sessionId, "power_point_delete_slide");

            Assert.Contains("\"isValid\":true", validateReorder, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"isValid\":true", validateDelete, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void BatchExecute_AcceptsPowerPointUnderscoreAliases()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "First", "Body");

            var batchPayload = service.BatchExecute(sessionId,
                "[{\"operation\":\"power_point_set_slide_title\",\"slideIndex\":1,\"title\":\"Updated\"}]");

            Assert.Contains("\"executed\":1", batchPayload, StringComparison.OrdinalIgnoreCase);
            var findResult = service.FindText(sessionId, "Updated");
            Assert.Contains("\"matchCount\":1", findResult, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ValidateOperation_AcceptsPowerPointAddSlideAlias()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            var payload = service.ValidateOperation(sessionId, "power_point_add_slide");

            Assert.Contains("\"isValid\":true", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void BatchExecute_FailurePayload_ContainsIndexAndErrorCode()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var payload = service.BatchExecute(sessionId,
                "[{\"operation\":\"word_append_paragraph\"},{\"operation\":\"word_append_paragraph\",\"text\":\"ok\"}]");

            Assert.Contains("\"failed\":1", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"index\":0", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"errorCode\":\"InvalidOperation\"", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"executed\":1", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ValidateOperation_AcceptsApplyStylePreset()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            var payload = service.ValidateOperation(sessionId, "apply_style_preset");

            Assert.Contains("\"isValid\":true", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ApplyStylePreset_PowerPointNeutral_UpdatesThemeName()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.ApplyStylePreset(sessionId, "neutral");
            service.CloseDocument(sessionId);

            using var presentation = PresentationDocument.Open(filePath, false);
            var themeName = presentation.PresentationPart?
                .SlideMasterParts
                .First()
                .ThemePart?
                .Theme?
                .Name?
                .Value;

            Assert.Equal("Neutral", themeName);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordSetParagraphStyle_AppliesFormatting()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Styled text");
            service.WordSetParagraphStyle(sessionId, 1, "Calibri", 20, bold: true, italic: true, colorHex: "FF0000");
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var runProperties = document.MainDocumentPart?
                .Document?
                .Body?
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .First()
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Run>()
                .First()
                .RunProperties;

            Assert.NotNull(runProperties);
            Assert.NotNull(runProperties!.Bold);
            Assert.NotNull(runProperties.Italic);
            Assert.Equal("FF0000", runProperties.Color?.Val?.Value);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void ExcelSetCellStyle_AppliesStyleIndex()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("xlsx");

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
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void PowerPointSetTextStyle_AndApplyTextPreset_Work()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Title", "Body");
            service.PowerPointSetTextStyle(sessionId, 1, 0, "Calibri", 30, bold: true, italic: false, colorHex: "AA0000");
            service.ApplyTextPreset(sessionId, "subtitle", 1);

            var payload = service.ValidateOperation(sessionId, "apply_text_preset");
            Assert.Contains("\"isValid\":true", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAddNumberedList_UsesDecimalDotByDefault()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddNumberedList(sessionId, "One\nTwo");
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var paragraph = document.MainDocumentPart?.Document?.Body?.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First();
            var numberingProps = paragraph?.ParagraphProperties?.NumberingProperties;

            Assert.NotNull(numberingProps);
            Assert.Equal(0, numberingProps!.NumberingLevelReference?.Val?.Value);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAddStructuredList_SupportsNestedMixedKinds()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var itemsJson = "[" +
                "{\"text\":\"Top 1\",\"level\":0,\"kind\":\"numbered\"}," +
                "{\"text\":\"Nested A\",\"level\":1,\"kind\":\"bulleted\",\"bulletStyle\":\"disc\"}," +
                "{\"text\":\"Nested B\",\"level\":1,\"kind\":\"bulleted\",\"bulletStyle\":\"square\"}" +
                "]";
            service.WordAddStructuredList(sessionId, itemsJson);
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var paragraphs = document.MainDocumentPart?.Document?.Body?.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList()
                ?? throw new InvalidOperationException("Paragraphs missing.");

            Assert.Equal(3, paragraphs.Count);
            Assert.Equal(0, paragraphs[0].ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value);
            Assert.Equal(1, paragraphs[1].ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value);
            Assert.Equal(1, paragraphs[2].ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAppendParagraph_AppliesDefaultSpacing()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Spacing test");
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var spacing = document.MainDocumentPart?
                .Document?
                .Body?
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .First()
                .ParagraphProperties?
                .SpacingBetweenLines;

            Assert.NotNull(spacing);
            Assert.Equal("160", spacing!.After?.Value);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordSetDocumentSpacingPreset_Comfortable_UpdatesParagraphSpacing()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "A");
            service.WordAppendParagraph(sessionId, "B");
            service.WordSetDocumentSpacingPreset(sessionId, "comfortable");
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var paragraphs = document.MainDocumentPart?
                .Document?
                .Body?
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .ToList() ?? [];

            Assert.NotEmpty(paragraphs);
            Assert.Equal("200", paragraphs[0].ParagraphProperties?.SpacingBetweenLines?.After?.Value);
            Assert.Equal("312", paragraphs[0].ParagraphProperties?.SpacingBetweenLines?.Line?.Value);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordInsertParagraphAfterText_InsertsAtExpectedPosition()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Anchor paragraph");
            var index = service.WordInsertParagraphAfterText(sessionId, "Anchor", "Inserted paragraph");

            Assert.Equal(2, index);
            var find = service.FindText(sessionId, "Inserted paragraph");
            Assert.Contains("\"index\":2", find, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordInsertTextAfterText_InsertsInline()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Hello world");
            var changed = service.WordInsertTextAfterText(sessionId, "Hello", " dear");

            Assert.True(changed);
            var find = service.FindText(sessionId, "Hello dear world");
            Assert.Contains("\"matchCount\":1", find, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordCreateOrUpdateStyle_AndApplyStyleByName_Work()
    {
        var service = new OfficeSessionService();
        var filePath = GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Styled paragraph");
            service.WordCreateOrUpdateStyle(sessionId, "MyBlueStyle", "{\"fontName\":\"Calibri\",\"fontSize\":14,\"colorHex\":\"1F4E79\",\"afterPt\":8,\"lineSpacing\":1.15}");
            service.WordApplyStyleByName(sessionId, 1, "MyBlueStyle");

            var styles = service.WordListStyles(sessionId);
            Assert.Contains("MyBlueStyle", styles, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            DeleteIfExists(filePath);
        }
    }
}
