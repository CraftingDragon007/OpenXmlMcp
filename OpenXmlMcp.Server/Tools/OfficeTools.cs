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
    public string CloseDocument(
        [Description("Active session id.")] string sessionId)
    {
        return officeSessionService.CloseDocument(sessionId);
    }

    [McpServerTool]
    [Description("Saves the active office session.")]
    public string SaveDocument(
        [Description("Active session id.")] string sessionId)
    {
        return officeSessionService.SaveDocument(sessionId);
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
    public string WordAppendParagraph(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Paragraph text to append.")] string text)
    {
        return officeSessionService.WordAppendParagraph(sessionId, text);
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
    [Description("Sets spacing for a paragraph by 1-based index in a DOCX session.")]
    public void WordSetParagraphSpacing(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("1-based paragraph index.")] int paragraphIndex,
        [Description("Spacing before in points.")] int beforePt,
        [Description("Spacing after in points.")] int afterPt,
        [Description("Line spacing multiplier, e.g. 1.15.")] double lineSpacing)
    {
        officeSessionService.WordSetParagraphSpacing(sessionId, paragraphIndex, beforePt, afterPt, lineSpacing);
    }

    [McpServerTool]
    [Description("Applies document-wide paragraph spacing preset in a DOCX session.")]
    public void WordSetDocumentSpacingPreset(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Preset: compact, normal, comfortable.")] string preset = "normal")
    {
        officeSessionService.WordSetDocumentSpacingPreset(sessionId, preset);
    }

    [McpServerTool]
    [Description("Inserts a paragraph after a paragraph containing anchor text.")]
    public int WordInsertParagraphAfterText(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Anchor text sequence to search.")] string anchorText,
        [Description("Paragraph text to insert.")] string text,
        [Description("1-based occurrence of the anchor text.")] int occurrence = 1,
        [Description("Use case-sensitive matching.")] bool matchCase = false)
    {
        return officeSessionService.WordInsertParagraphAfterText(sessionId, anchorText, text, occurrence, matchCase);
    }

    [McpServerTool]
    [Description("Inserts inline text immediately after a text sequence match.")]
    public bool WordInsertTextAfterText(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Anchor text sequence to search.")] string anchorText,
        [Description("Text to insert after the anchor.")] string text,
        [Description("1-based occurrence of the anchor text.")] int occurrence = 1,
        [Description("Use case-sensitive matching.")] bool matchCase = false)
    {
        return officeSessionService.WordInsertTextAfterText(sessionId, anchorText, text, occurrence, matchCase);
    }

    [McpServerTool]
    [Description("Lists available paragraph/character styles in a DOCX session.")]
    public string WordListStyles(
        [Description("Active DOCX session id.")] string sessionId)
    {
        return officeSessionService.WordListStyles(sessionId);
    }

    [McpServerTool]
    [Description("Applies an existing Word style by style id or style name.")]
    public void WordApplyStyleByName(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("1-based paragraph index.")] int paragraphIndex,
        [Description("Style id or style name.")] string styleName)
    {
        officeSessionService.WordApplyStyleByName(sessionId, paragraphIndex, styleName);
    }

    [McpServerTool]
    [Description("Creates or updates a custom Word paragraph style from JSON options.")]
    public void WordCreateOrUpdateStyle(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Custom style display name.")] string styleName,
        [Description("JSON options for style properties.")] string styleJson)
    {
        officeSessionService.WordCreateOrUpdateStyle(sessionId, styleName, styleJson);
    }

    [McpServerTool]
    [Description("Inserts a paragraph at a 1-based index in a DOCX session.")]
    public string WordInsertParagraphAt(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("1-based paragraph insertion index.")] int index,
        [Description("Paragraph text.")] string text)
    {
        return officeSessionService.WordInsertParagraphAt(sessionId, index, text);
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
    public string WordAddBulletedList(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Newline-separated bullet lines.")] string lines,
        [Description("Bullet style (default: disc).")]
        string bulletStyle = "disc")
    {
        return officeSessionService.WordAddBulletedList(sessionId, lines, bulletStyle);
    }

    [McpServerTool]
    [Description("Adds a numbered list to a DOCX session using newline-separated items.")]
    public string WordAddNumberedList(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Newline-separated numbered lines.")] string lines,
        [Description("Number style (default: decimal-dot).")]
        string numberStyle = "decimal-dot")
    {
        return officeSessionService.WordAddNumberedList(sessionId, lines, numberStyle);
    }

    [McpServerTool]
    [Description("Adds a structured mixed list (numbered/bulleted, nested) from JSON items.")]
    public string WordAddStructuredList(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("JSON array of items: text, level, kind, optional bulletStyle/numberStyle.")]
        string itemsJson)
    {
        return officeSessionService.WordAddStructuredList(sessionId, itemsJson);
    }

    [McpServerTool]
    [Description("Adds a plain table to a DOCX session.")]
    public string WordAddTable(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("Number of rows.")] int rows,
        [Description("Number of columns.")] int columns)
    {
        return officeSessionService.WordAddTable(sessionId, rows, columns);
    }

    [McpServerTool]
    [Description("Sets text in a Word table cell by 1-based table/row/column indexes.")]
    public string WordSetTableCell(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("1-based table index in document body.")] int tableIndex,
        [Description("1-based row index within the table.")] int rowIndex,
        [Description("1-based column index within the row.")] int columnIndex,
        [Description("Cell text value.")] string text)
    {
        return officeSessionService.WordSetTableCell(sessionId, tableIndex, rowIndex, columnIndex, text);
    }

    [McpServerTool]
    [Description("Gets text from a Word table cell by 1-based table/row/column indexes.")]
    public string WordGetTableCell(
        [Description("Active DOCX session id.")] string sessionId,
        [Description("1-based table index in document body.")] int tableIndex,
        [Description("1-based row index within the table.")] int rowIndex,
        [Description("1-based column index within the row.")] int columnIndex)
    {
        return officeSessionService.WordGetTableCell(sessionId, tableIndex, rowIndex, columnIndex);
    }

    [McpServerTool]
    [Description("Sets a string value in an XLSX cell.")]
    public string ExcelSetCellValue(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName,
        [Description("Cell reference, e.g. A1.")] string cellReference,
        [Description("Cell value.")] string value)
    {
        return officeSessionService.ExcelSetCellValue(sessionId, sheetName, cellReference, value);
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
    [Description("Gets detailed cell info from an XLSX session, including formula and cached value.")]
    public string ExcelGetCellInfo(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName,
        [Description("Cell reference, e.g. A1.")] string cellReference)
    {
        return officeSessionService.ExcelGetCellInfo(sessionId, sheetName, cellReference);
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
    [Description("Sets a 2D range of values in an XLSX sheet. valuesJson must be strict JSON array-of-arrays (double-quoted strings). Strings starting with '=' are stored as formulas.")]
    public string ExcelSetRangeValues(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName,
        [Description("Top-left start cell, e.g. A1.")] string startCell,
        [Description("JSON array of arrays for values.")] string valuesJson)
    {
        return officeSessionService.ExcelSetRangeValues(sessionId, sheetName, startCell, valuesJson);
    }

    [McpServerTool]
    [Description("Sets a formula in an XLSX cell.")]
    public string ExcelSetFormula(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("Sheet name, e.g. Sheet1.")] string sheetName,
        [Description("Cell reference, e.g. C1.")] string cellReference,
        [Description("Formula with or without leading '='.")] string formula)
    {
        return officeSessionService.ExcelSetFormula(sessionId, sheetName, cellReference, formula);
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
    public string ExcelAddWorksheet(
        [Description("Active XLSX session id.")] string sessionId,
        [Description("New sheet name.")] string sheetName)
    {
        return officeSessionService.ExcelAddWorksheet(sessionId, sheetName);
    }

    [McpServerTool]
    [Description("Adds a slide to a PPTX session. bodyType can be text, bulleted, or numbered.")]
    public string PowerPointAddSlide(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("Slide title.")] string title,
        [Description("Slide body text.")] string body,
        [Description("Body type: text, bulleted, numbered.")] string bodyType = "text")
    {
        return officeSessionService.PowerPointAddSlide(sessionId, title, body, bodyType);
    }

    [McpServerTool]
    [Description("Inserts a slide at a 1-based index in a PPTX session.")]
    public string PowerPointInsertSlideAt(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based insertion index.")] int index,
        [Description("Slide title.")] string title,
        [Description("Slide body text.")] string body)
    {
        return officeSessionService.PowerPointInsertSlideAt(sessionId, index, title, body);
    }

    [McpServerTool]
    [Description("Sets the title text for a slide by 1-based index.")]
    public string PowerPointSetSlideTitle(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based slide index.")] int slideIndex,
        [Description("New title text.")] string title)
    {
        return officeSessionService.PowerPointSetSlideTitle(sessionId, slideIndex, title);
    }

    [McpServerTool]
    [Description("Sets the body content for a slide by 1-based index. bodyType can be text, bulleted, or numbered.")]
    public string PowerPointSetSlideBody(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based slide index.")] int slideIndex,
        [Description("New body text.")] string body,
        [Description("Body type: text, bulleted, numbered.")] string bodyType = "text")
    {
        return officeSessionService.PowerPointSetSlideBody(sessionId, slideIndex, body, bodyType);
    }

    [McpServerTool]
    [Description("Sets hidden speaker notes for a slide by 1-based index.")]
    public string PowerPointSetSlideNotes(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based slide index.")] int slideIndex,
        [Description("Speaker notes text.")] string notes)
    {
        return officeSessionService.PowerPointSetSlideNotes(sessionId, slideIndex, notes);
    }

    [McpServerTool]
    [Description("Gets hidden speaker notes for a slide by 1-based index.")]
    public string PowerPointGetSlideNotes(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based slide index.")] int slideIndex)
    {
        return officeSessionService.PowerPointGetSlideNotes(sessionId, slideIndex);
    }

    [McpServerTool]
    [Description("Applies font style to a text slot on a slide (slot 0=title, 1=entire body across all body paragraphs).")]
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
    public string PowerPointReorderSlide(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based source slide index.")] int fromIndex,
        [Description("1-based destination slide index.")] int toIndex)
    {
        return officeSessionService.PowerPointReorderSlide(sessionId, fromIndex, toIndex);
    }

    [McpServerTool]
    [Description("Deletes a slide by 1-based index.")]
    public string PowerPointDeleteSlide(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("1-based slide index.")] int slideIndex)
    {
        return officeSessionService.PowerPointDeleteSlide(sessionId, slideIndex);
    }

    [McpServerTool]
    [Description("Adds a bullet style slide to a PPTX session. Provide newline-separated bullet text.")]
    public string PowerPointAddBulletSlide(
        [Description("Active PPTX session id.")] string sessionId,
        [Description("Slide title.")] string title,
        [Description("Newline-separated bullet lines.")] string bulletLines)
    {
        return officeSessionService.PowerPointAddBulletSlide(sessionId, title, bulletLines);
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
    public string ApplyStylePreset(
        [Description("Active session id.")] string sessionId,
        [Description("Style preset name, e.g. default or neutral.")] string preset = "default")
    {
        return officeSessionService.ApplyStylePreset(sessionId, preset);
    }

    [McpServerTool]
    [Description("Lists available style presets for the active document type.")]
    public string ListStylePresets(
        [Description("Active session id.")] string sessionId)
    {
        return officeSessionService.ListStylePresets(sessionId);
    }

    [McpServerTool]
    [Description("Applies text presets such as title or subtitle.")]
    public string ApplyTextPreset(
        [Description("Active session id.")] string sessionId,
        [Description("Text preset: default, title, subtitle.")] string preset,
        [Description("Target index: paragraph/slide index depending on document type.")] int targetIndex = 1)
    {
        return officeSessionService.ApplyTextPreset(sessionId, preset, targetIndex);
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
    public string UndoLastChange(
        [Description("Active session id.")] string sessionId)
    {
        return officeSessionService.UndoLastChange(sessionId);
    }
}
