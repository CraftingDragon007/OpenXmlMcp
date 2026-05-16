using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlMcp.Server.Services;

namespace OpenXmlMcp.Server.Tests;

public partial class OfficeSessionServiceTests
{
    [Fact]
    public void GetDocumentInfo_ReturnsSessionPayload()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

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
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ValidateOperation_ReadOnlySession_ReturnsInvalidForWriteOperation()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

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
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void BatchExecute_MixedOperations_ReportsFailuresAndSuccesses()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

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
            Assert.Contains("\"results\"", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void BatchExecute_AcceptsWrappedPayloadAndOperationNameAlias()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var operations = "{" +
                "\"operations\":[" +
                "{\"operationName\":\"word_append_paragraph\",\"text\":\"hello\"}" +
                "]}";

            var payload = service.BatchExecute(sessionId, operations);
            Assert.Contains("\"executed\":1", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"failed\":0", payload, StringComparison.OrdinalIgnoreCase);

            var find = service.FindText(sessionId, "hello");
            Assert.Contains("\"matchCount\":1", find, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void BatchExecute_InvalidPayload_ReturnsHelpfulErrorCode()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");

            var ex = Assert.Throws<InvalidOperationException>(() => service.BatchExecute(sessionId, "{\"foo\":1}"));
            Assert.Contains("Invalid batch payload", ex.Message, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ValidateOperation_KnowsPhase2OperationNames()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var payload = service.ValidateOperation(sessionId, "word_add_table");

            Assert.Contains("\"isValid\":true", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void UndoLastChange_RestoresPreviousState()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

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
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void GetOperationHistory_ReturnsEntries()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

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
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void OpenDocument_LargeFile_ThrowsSafetyException()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
            stream.SetLength(21L * 1024 * 1024);

            var ex = Assert.Throws<InvalidOperationException>(() => service.OpenDocument(filePath));
            Assert.Contains("exceeds safety limit", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ValidateOperation_AcceptsPowerPointUnderscoreAliases()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

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
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void BatchExecute_AcceptsPowerPointUnderscoreAliases()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

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
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ValidateOperation_AcceptsPowerPointAddSlideAlias()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            var payload = service.ValidateOperation(sessionId, "power_point_add_slide");

            Assert.Contains("\"isValid\":true", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void BatchExecute_FailurePayload_ContainsIndexAndErrorCode()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var payload = service.BatchExecute(sessionId,
                "[{\"operation\":\"word_append_paragraph\"},{\"operation\":\"word_append_paragraph\",\"text\":\"ok\"}]");

            Assert.Contains("\"failed\":1", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"index\":0", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"errorCode\":\"MissingField\"", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"executed\":1", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ValidateOperation_AcceptsApplyStylePreset()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("xlsx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "xlsx");
            var payload = service.ValidateOperation(sessionId, "apply_style_preset");

            Assert.Contains("\"isValid\":true", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ApplyStylePreset_PowerPointNeutral_UpdatesThemeName()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

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
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ListStylePresets_ReturnsDocumentTypeSpecificPresets()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            var payload = service.ListStylePresets(sessionId);

            Assert.Contains("PowerPoint", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("neutral", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void ApplyStylePreset_InvalidPreset_ThrowsHelpfulMessage()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var ex = Assert.Throws<InvalidOperationException>(() => service.ApplyStylePreset(sessionId, "neutral"));

            Assert.Contains("Invalid preset", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("default", ex.Message, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void MutationOperations_ReturnStructuredPayload()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var payload = service.WordAppendParagraph(sessionId, "Payload check");

            using var doc = JsonDocument.Parse(payload);
            Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());
            Assert.Equal("word_append_paragraph", doc.RootElement.GetProperty("operation").GetString());
            Assert.True(doc.RootElement.GetProperty("changed").GetBoolean());

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
}
