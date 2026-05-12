using System.ComponentModel;
using ModelContextProtocol.Server;
using OpenXmlMcp.Server.Services;

namespace OpenXmlMcp.Server.Tools;

internal class OfficeTools(OfficeSessionService officeSessionService)
{
    [McpServerTool]
    [Description("Opens an existing Office document and returns a session id.")]
    public string OpenDocument(
        [Description("Absolute or relative path to a .docx/.xlsx/.pptx file.")] string filePath,
        [Description("Open in read-only mode.")] bool readOnly = false)
    {
        return officeSessionService.OpenDocument(filePath, readOnly);
    }

    [McpServerTool]
    [Description("Creates a new Office document and returns a session id.")]
    public string CreateDocument(
        [Description("Absolute or relative path for the new file.")] string filePath,
        [Description("Document type: docx, xlsx, or pptx.")] string documentType)
    {
        return officeSessionService.CreateDocument(filePath, documentType);
    }

    [McpServerTool]
    [Description("Closes an open office session.")]
    public void CloseDocument(
        [Description("Active session id.")] string sessionId)
    {
        officeSessionService.CloseDocument(sessionId);
    }

    [McpServerTool]
    [Description("Saves the active office session.")]
    public void SaveDocument(
        [Description("Active session id.")] string sessionId)
    {
        officeSessionService.SaveDocument(sessionId);
    }

    [McpServerTool]
    [Description("Returns metadata for an open office session.")]
    public string GetDocumentInfo(
        [Description("Active session id.")] string sessionId)
    {
        return officeSessionService.GetDocumentInfo(sessionId);
    }

    [McpServerTool]
    [Description("Lists document structure for an open office session.")]
    public string ListStructure(
        [Description("Active session id.")] string sessionId)
    {
        return officeSessionService.ListStructure(sessionId);
    }

    [McpServerTool]
    [Description("Finds text matches in an open office session.")]
    public string FindText(
        [Description("Active session id.")] string sessionId,
        [Description("Case-insensitive search text.")] string query)
    {
        return officeSessionService.FindText(sessionId, query);
    }

    [McpServerTool]
    [Description("Validates whether an operation is safe for the current session.")]
    public string ValidateOperation(
        [Description("Active session id.")] string sessionId,
        [Description("Operation name, e.g. word_append_paragraph.")] string operationName)
    {
        return officeSessionService.ValidateOperation(sessionId, operationName);
    }

    [McpServerTool]
    [Description("Appends a paragraph to a DOCX session.")]
    public void WordAppendParagraph(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Paragraph text to append.")] string text)
    {
        officeSessionService.WordAppendParagraph(sessionId, text);
    }

    [McpServerTool]
    [Description("Sets a string value in an XLSX cell.")]
    public void ExcelSetCellValue(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName,
        [Description("Cell reference, e.g. A1.")] string cellReference,
        [Description("Cell value.")] string value)
    {
        officeSessionService.ExcelSetCellValue(sessionId, sheetName, cellReference, value);
    }

    [McpServerTool]
    [Description("Gets a cell value from an XLSX session.")]
    public string ExcelGetCellValue(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName,
        [Description("Cell reference, e.g. A1.")] string cellReference)
    {
        return officeSessionService.ExcelGetCellValue(sessionId, sheetName, cellReference);
    }

    [McpServerTool]
    [Description("Adds a simple title/body slide to a PPTX session.")]
    public void PowerPointAddSlide(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("Slide title.")] string title,
        [Description("Slide body text.")] string body)
    {
        officeSessionService.PowerPointAddSlide(sessionId, title, body);
    }
}
