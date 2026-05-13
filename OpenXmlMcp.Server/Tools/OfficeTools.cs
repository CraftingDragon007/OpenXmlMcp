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
    [Description("Applies font style to a paragraph by 1-based index in a DOCX session.")]
    public void WordSetParagraphStyle(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("1-based paragraph index.")] int paragraphIndex,
        [Description("Font family name.")] string fontName,
        [Description("Font size in points (8-96).")]
        int fontSize,
        [Description("Whether text is bold.")] bool bold,
        [Description("Whether text is italic.")] bool italic,
        [Description("Hex text color without '#', e.g. FF0000.")] string colorHex)
    {
        officeSessionService.WordSetParagraphStyle(sessionId, paragraphIndex, fontName, fontSize, bold, italic, colorHex);
    }

    [McpServerTool]
    [Description("Inserts a paragraph at a 1-based index in a DOCX session.")]
    public void WordInsertParagraphAt(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("1-based paragraph insertion index.")] int index,
        [Description("Paragraph text.")] string text)
    {
        officeSessionService.WordInsertParagraphAt(sessionId, index, text);
    }

    [McpServerTool]
    [Description("Replaces text in a DOCX session and returns replacement count.")]
    public int WordReplaceText(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Text to find.")] string find,
        [Description("Replacement text.")] string replace,
        [Description("Use case-sensitive matching.")] bool matchCase = false)
    {
        return officeSessionService.WordReplaceText(sessionId, find, replace, matchCase);
    }

    [McpServerTool]
    [Description("Adds a heading paragraph to a DOCX session.")]
    public void WordAddHeading(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Heading level from 1 to 6.")] int level,
        [Description("Heading text.")] string text)
    {
        officeSessionService.WordAddHeading(sessionId, level, text);
    }

    [McpServerTool]
    [Description("Adds a bulleted list to a DOCX session using newline-separated items.")]
    public void WordAddBulletedList(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Newline-separated bullet lines.")] string lines,
        [Description("Bullet style (default: disc).")]
        string bulletStyle = "disc")
    {
        officeSessionService.WordAddBulletedList(sessionId, lines, bulletStyle);
    }

    [McpServerTool]
    [Description("Adds a numbered list to a DOCX session using newline-separated items.")]
    public void WordAddNumberedList(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Newline-separated numbered lines.")] string lines,
        [Description("Number style (default: decimal-dot).")]
        string numberStyle = "decimal-dot")
    {
        officeSessionService.WordAddNumberedList(sessionId, lines, numberStyle);
    }

    [McpServerTool]
    [Description("Adds a structured mixed list (numbered/bulleted, nested) from JSON items.")]
    public void WordAddStructuredList(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("JSON array of items: text, level, kind, optional bulletStyle/numberStyle.")]
        string itemsJson)
    {
        officeSessionService.WordAddStructuredList(sessionId, itemsJson);
    }

    [McpServerTool]
    [Description("Adds a plain table to a DOCX session.")]
    public void WordAddTable(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Number of rows.")] int rows,
        [Description("Number of columns.")] int columns)
    {
        officeSessionService.WordAddTable(sessionId, rows, columns);
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
    [Description("Applies font style to an XLSX cell.")]
    public void ExcelSetCellStyle(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName,
        [Description("Cell reference, e.g. A1.")] string cellReference,
        [Description("Font family name.")] string fontName,
        [Description("Font size in points (8-96).")]
        int fontSize,
        [Description("Whether text is bold.")] bool bold,
        [Description("Whether text is italic.")] bool italic,
        [Description("Hex text color without '#', e.g. FF0000.")] string colorHex)
    {
        officeSessionService.ExcelSetCellStyle(sessionId, sheetName, cellReference, fontName, fontSize, bold, italic, colorHex);
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
    [Description("Returns the used range bounds for an XLSX worksheet.")]
    public string ExcelGetUsedRange(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName)
    {
        return officeSessionService.ExcelGetUsedRange(sessionId, sheetName);
    }

    [McpServerTool]
    [Description("Sets a 2D range of values in an XLSX sheet. valuesJson must be a JSON array of arrays.")]
    public void ExcelSetRangeValues(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName,
        [Description("Top-left start cell, e.g. A1.")] string startCell,
        [Description("JSON array of arrays for values.")] string valuesJson)
    {
        officeSessionService.ExcelSetRangeValues(sessionId, sheetName, startCell, valuesJson);
    }

    [McpServerTool]
    [Description("Sets a formula in an XLSX cell.")]
    public void ExcelSetFormula(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName,
        [Description("Cell reference, e.g. C1.")] string cellReference,
        [Description("Formula with or without leading '='.")] string formula)
    {
        officeSessionService.ExcelSetFormula(sessionId, sheetName, cellReference, formula);
    }

    [McpServerTool]
    [Description("Gets a formula from an XLSX cell.")]
    public string ExcelGetFormula(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName,
        [Description("Cell reference, e.g. C1.")] string cellReference)
    {
        return officeSessionService.ExcelGetFormula(sessionId, sheetName, cellReference);
    }

    [McpServerTool]
    [Description("Adds a new worksheet to an XLSX session.")]
    public void ExcelAddWorksheet(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("New sheet name.")] string sheetName)
    {
        officeSessionService.ExcelAddWorksheet(sessionId, sheetName);
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

    [McpServerTool]
    [Description("Inserts a slide at a 1-based index in a PPTX session.")]
    public void PowerPointInsertSlideAt(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based insertion index.")] int index,
        [Description("Slide title.")] string title,
        [Description("Slide body text.")] string body)
    {
        officeSessionService.PowerPointInsertSlideAt(sessionId, index, title, body);
    }

    [McpServerTool]
    [Description("Sets the title text for a slide by 1-based index.")]
    public void PowerPointSetSlideTitle(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based slide index.")] int slideIndex,
        [Description("New title text.")] string title)
    {
        officeSessionService.PowerPointSetSlideTitle(sessionId, slideIndex, title);
    }

    [McpServerTool]
    [Description("Sets the body text for a slide by 1-based index.")]
    public void PowerPointSetSlideBody(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based slide index.")] int slideIndex,
        [Description("New body text.")] string body)
    {
        officeSessionService.PowerPointSetSlideBody(sessionId, slideIndex, body);
    }

    [McpServerTool]
    [Description("Applies font style to a text slot on a slide (slot 0=title, 1=body).")]
    public void PowerPointSetTextStyle(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based slide index.")] int slideIndex,
        [Description("Text slot index (0=title, 1=body).")]
        int slot,
        [Description("Font family name.")] string fontName,
        [Description("Font size in points (8-96).")]
        int fontSize,
        [Description("Whether text is bold.")] bool bold,
        [Description("Whether text is italic.")] bool italic,
        [Description("Hex text color without '#', e.g. FF0000.")] string colorHex)
    {
        officeSessionService.PowerPointSetTextStyle(sessionId, slideIndex, slot, fontName, fontSize, bold, italic, colorHex);
    }

    [McpServerTool]
    [Description("Moves a slide from one 1-based index to another.")]
    public void PowerPointReorderSlide(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based source slide index.")] int fromIndex,
        [Description("1-based destination slide index.")] int toIndex)
    {
        officeSessionService.PowerPointReorderSlide(sessionId, fromIndex, toIndex);
    }

    [McpServerTool]
    [Description("Deletes a slide by 1-based index.")]
    public void PowerPointDeleteSlide(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based slide index.")] int slideIndex)
    {
        officeSessionService.PowerPointDeleteSlide(sessionId, slideIndex);
    }

    [McpServerTool]
    [Description("Adds a bullet style slide to a PPTX session. Provide newline-separated bullet text.")]
    public void PowerPointAddBulletSlide(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("Slide title.")] string title,
        [Description("Newline-separated bullet lines.")] string bulletLines)
    {
        officeSessionService.PowerPointAddBulletSlide(sessionId, title, bulletLines);
    }

    [McpServerTool]
    [Description("Executes multiple operations for one active session. Input is a JSON array of operation objects.")]
    public string BatchExecute(
        [Description("Active session id.")] string sessionId,
        [Description("JSON array of operations with operation-specific fields.")] string operationsJson)
    {
        return officeSessionService.BatchExecute(sessionId, operationsJson);
    }

    [McpServerTool]
    [Description("Applies a cross-suite style preset to the active document session.")]
    public void ApplyStylePreset(
        [Description("Active session id.")] string sessionId,
        [Description("Style preset name, e.g. default or neutral.")] string preset = "default")
    {
        officeSessionService.ApplyStylePreset(sessionId, preset);
    }

    [McpServerTool]
    [Description("Applies text presets such as title or subtitle.")]
    public void ApplyTextPreset(
        [Description("Active session id.")] string sessionId,
        [Description("Text preset: default, title, subtitle.")] string preset,
        [Description("Target index: paragraph/slide index depending on document type.")] int targetIndex = 1)
    {
        officeSessionService.ApplyTextPreset(sessionId, preset, targetIndex);
    }

    [McpServerTool]
    [Description("Returns operation history entries for a session.")]
    public string GetOperationHistory(
        [Description("Active session id.")] string sessionId)
    {
        return officeSessionService.GetOperationHistory(sessionId);
    }

    [McpServerTool]
    [Description("Restores the previous checkpoint for the active session.")]
    public void UndoLastChange(
        [Description("Active session id.")] string sessionId)
    {
        officeSessionService.UndoLastChange(sessionId);
    }
}
