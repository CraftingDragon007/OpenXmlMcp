using System.Text.Json.Nodes;

namespace OpenXmlMcp.Server.Services;

internal static class OfficeBatchDispatcher
{
    public delegate void BatchOperationHandler(OfficeSessionService service, string sessionId, JsonNode payload);

    private static readonly IReadOnlyDictionary<string, BatchOperationHandler> Handlers = BuildHandlers();

    public static void Dispatch(OfficeSessionService service, string sessionId, string operation, JsonNode payload)
    {
        var normalized = operation.Trim().ToLowerInvariant();
        var spec = OfficeOperationRegistry.TryGet(normalized) ?? throw new InvalidOperationException($"Unsupported batch operation '{operation}'.");
        if (!Handlers.TryGetValue(spec.CanonicalName, out var handler))
        {
            throw new InvalidOperationException($"Operation '{operation}' is not supported in batch execution.");
        }

        handler(service, sessionId, payload);
    }

    private static IReadOnlyDictionary<string, BatchOperationHandler> BuildHandlers()
    {
        return new Dictionary<string, BatchOperationHandler>(StringComparer.OrdinalIgnoreCase)
        {
            ["word_append_paragraph"] = static (svc, sessionId, payload) =>
                _ = svc.WordAppendParagraph(sessionId, payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'.")),
            ["word_set_paragraph_style"] = static (svc, sessionId, payload) =>
                _ = svc.WordSetParagraphStyle(
                    sessionId,
                    payload["paragraphIndex"]?.GetValue<int>() ?? 1,
                    payload["fontName"]?.GetValue<string>() ?? "Calibri",
                    payload["fontSize"]?.GetValue<int>() ?? 18,
                    payload["bold"]?.GetValue<bool>() ?? false,
                    payload["italic"]?.GetValue<bool>() ?? false,
                    payload["colorHex"]?.GetValue<string>() ?? "000000"),
            ["word_apply_table_style"] = static (svc, sessionId, payload) =>
                _ = svc.WordApplyTableStyle(
                    sessionId,
                    payload["tableIndex"]?.GetValue<int>() ?? throw new InvalidOperationException("Missing 'tableIndex'."),
                    payload["styleName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'styleName'.")),
            ["word_format_table_header_row"] = static (svc, sessionId, payload) =>
                _ = svc.WordFormatTableHeaderRow(
                    sessionId,
                    payload["tableIndex"]?.GetValue<int>() ?? throw new InvalidOperationException("Missing 'tableIndex'."),
                    payload["bold"]?.GetValue<bool>() ?? true,
                    payload["shadingFill"]?.GetValue<string>(),
                    payload["fontName"]?.GetValue<string>(),
                    payload["colorHex"]?.GetValue<string>()),
            ["word_set_table_values"] = static (svc, sessionId, payload) =>
                _ = svc.WordSetTableValues(
                    sessionId,
                    payload["tableIndex"]?.GetValue<int>() ?? throw new InvalidOperationException("Missing 'tableIndex'."),
                    payload["valuesJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'valuesJson'."),
                    payload["startRow"]?.GetValue<int>() ?? 1,
                    payload["startColumn"]?.GetValue<int>() ?? 1),
            ["word_add_table"] = static (svc, sessionId, payload) =>
                _ = svc.WordAddTable(sessionId, payload["rows"]?.GetValue<int>() ?? 0, payload["columns"]?.GetValue<int>() ?? 0),
            ["word_set_table_cell"] = static (svc, sessionId, payload) =>
                _ = svc.WordSetTableCell(
                    sessionId,
                    payload["tableIndex"]?.GetValue<int>() ?? throw new InvalidOperationException("Missing 'tableIndex'."),
                    payload["rowIndex"]?.GetValue<int>() ?? throw new InvalidOperationException("Missing 'rowIndex'."),
                    payload["columnIndex"]?.GetValue<int>() ?? throw new InvalidOperationException("Missing 'columnIndex'."),
                    payload["text"]?.GetValue<string>() ?? string.Empty),
            ["word_get_table_cell"] = static (svc, sessionId, payload) =>
                _ = svc.WordGetTableCell(
                    sessionId,
                    payload["tableIndex"]?.GetValue<int>() ?? throw new InvalidOperationException("Missing 'tableIndex'."),
                    payload["rowIndex"]?.GetValue<int>() ?? throw new InvalidOperationException("Missing 'rowIndex'."),
                    payload["columnIndex"]?.GetValue<int>() ?? throw new InvalidOperationException("Missing 'columnIndex'.")),
            ["word_insert_paragraph_at"] = static (svc, sessionId, payload) =>
                _ = svc.WordInsertParagraphAt(sessionId, payload["index"]?.GetValue<int>() ?? 0, payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'.")),
            ["word_replace_text"] = static (svc, sessionId, payload) =>
                _ = svc.WordReplaceText(sessionId, payload["find"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'find'."), payload["replace"]?.GetValue<string>() ?? string.Empty, payload["matchCase"]?.GetValue<bool>() ?? false),
            ["word_add_heading"] = static (svc, sessionId, payload) =>
                _ = svc.WordAddHeading(sessionId, payload["level"]?.GetValue<int>() ?? 1, payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'.")),
            ["word_add_bulleted_list"] = static (svc, sessionId, payload) =>
                _ = svc.WordAddBulletedList(
                    sessionId,
                    payload["lines"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'lines'."),
                    payload["bulletStyle"]?.GetValue<string>() ?? "disc"),
            ["word_add_numbered_list"] = static (svc, sessionId, payload) =>
                _ = svc.WordAddNumberedList(
                    sessionId,
                    payload["lines"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'lines'."),
                    payload["numberStyle"]?.GetValue<string>() ?? "decimal-dot"),
            ["word_add_structured_list"] = static (svc, sessionId, payload) =>
                _ = svc.WordAddStructuredList(
                    sessionId,
                    payload["itemsJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'itemsJson'.")),
            ["word_set_paragraph_spacing"] = static (svc, sessionId, payload) =>
                _ = svc.WordSetParagraphSpacing(
                    sessionId,
                    payload["paragraphIndex"]?.GetValue<int>() ?? 1,
                    payload["beforePt"]?.GetValue<int>() ?? 0,
                    payload["afterPt"]?.GetValue<int>() ?? 8,
                    payload["lineSpacing"]?.GetValue<double>() ?? 1.15),
            ["word_set_document_spacing_preset"] = static (svc, sessionId, payload) =>
                _ = svc.WordSetDocumentSpacingPreset(
                    sessionId,
                    payload["preset"]?.GetValue<string>() ?? "normal"),
            ["word_insert_paragraph_after_text"] = static (svc, sessionId, payload) =>
                _ = svc.WordInsertParagraphAfterText(
                    sessionId,
                    payload["anchorText"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'anchorText'."),
                    payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'."),
                    payload["occurrence"]?.GetValue<int>() ?? 1,
                    payload["matchCase"]?.GetValue<bool>() ?? false),
            ["word_insert_text_after_text"] = static (svc, sessionId, payload) =>
                _ = svc.WordInsertTextAfterText(
                    sessionId,
                    payload["anchorText"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'anchorText'."),
                    payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'."),
                    payload["occurrence"]?.GetValue<int>() ?? 1,
                    payload["matchCase"]?.GetValue<bool>() ?? false),
            ["word_apply_style_by_name"] = static (svc, sessionId, payload) =>
                _ = svc.WordApplyStyleByName(
                    sessionId,
                    payload["paragraphIndex"]?.GetValue<int>() ?? 1,
                    payload["styleName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'styleName'.")),
            ["word_create_or_update_style"] = static (svc, sessionId, payload) =>
                _ = svc.WordCreateOrUpdateStyle(
                    sessionId,
                    payload["styleName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'styleName'."),
                    payload["styleJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'styleJson'.")),
            ["word_apply_character_style_to_all"] = static (svc, sessionId, payload) =>
                _ = svc.WordApplyCharacterStyleToAll(
                    sessionId,
                    payload["queriesJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'queriesJson'."),
                    payload["styleName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'styleName'."),
                    payload["matchCase"]?.GetValue<bool>() ?? false,
                    payload["wholeWord"]?.GetValue<bool>() ?? true),
            ["word_apply_character_style_by_pattern"] = static (svc, sessionId, payload) =>
                _ = svc.WordApplyCharacterStyleByPattern(
                    sessionId,
                    payload["pattern"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'pattern'."),
                    payload["styleName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'styleName'."),
                    payload["matchCase"]?.GetValue<bool>() ?? true,
                    payload["maxMatches"]?.GetValue<int>() ?? 5000),
            ["word_insert_after_heading"] = static (svc, sessionId, payload) =>
                _ = svc.WordInsertAfterHeading(
                    sessionId,
                    payload["headingText"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'headingText'."),
                    payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'."),
                    payload["occurrence"]?.GetValue<int>() ?? 1,
                    payload["matchCase"]?.GetValue<bool>() ?? false),
            ["word_replace_section"] = static (svc, sessionId, payload) =>
                _ = svc.WordReplaceSection(
                    sessionId,
                    payload["headingText"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'headingText'."),
                    payload["replacementJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'replacementJson'."),
                    payload["occurrence"]?.GetValue<int>() ?? 1,
                    payload["matchCase"]?.GetValue<bool>() ?? false),
            ["word_apply_character_style_to_text"] = static (svc, sessionId, payload) =>
                _ = svc.WordApplyCharacterStyleToText(
                    sessionId,
                    payload["anchorText"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'anchorText'."),
                    payload["styleName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'styleName'."),
                    payload["occurrence"]?.GetValue<int>() ?? 1,
                    payload["matchCase"]?.GetValue<bool>() ?? false),
            ["word_list_paragraph_runs"] = static (svc, sessionId, payload) =>
                _ = svc.WordListParagraphRuns(
                    sessionId,
                    payload["paragraphIndex"]?.GetValue<int>() ?? 1),
            ["word_get_paragraph_info"] = static (svc, sessionId, payload) =>
                _ = svc.WordGetParagraphInfo(
                    sessionId,
                    payload["paragraphIndex"]?.GetValue<int>() ?? 1),
            ["word_list_styles"] = static (svc, sessionId, _) => _ = svc.WordListStyles(sessionId),
            ["excel_set_cell_value"] = static (svc, sessionId, payload) =>
                _ = svc.ExcelSetCellValue(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."), payload["cellReference"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'cellReference'."), payload["value"]?.GetValue<string>() ?? string.Empty),
            ["excel_set_cell_style"] = static (svc, sessionId, payload) =>
                _ = svc.ExcelSetCellStyle(
                    sessionId,
                    payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."),
                    payload["cellReference"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'cellReference'."),
                    payload["fontName"]?.GetValue<string>() ?? "Calibri",
                    payload["fontSize"]?.GetValue<int>() ?? 11,
                    payload["bold"]?.GetValue<bool>() ?? false,
                    payload["italic"]?.GetValue<bool>() ?? false,
                    payload["colorHex"]?.GetValue<string>() ?? "000000"),
            ["excel_set_range_values"] = static (svc, sessionId, payload) =>
                _ = svc.ExcelSetRangeValues(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."), payload["startCell"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'startCell'."), payload["valuesJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'valuesJson'.")),
            ["excel_set_formula"] = static (svc, sessionId, payload) =>
                _ = svc.ExcelSetFormula(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."), payload["cellReference"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'cellReference'."), payload["formula"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'formula'.")),
            ["excel_add_worksheet"] = static (svc, sessionId, payload) =>
                _ = svc.ExcelAddWorksheet(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'.")),
            ["powerpoint_add_slide"] = static (svc, sessionId, payload) =>
                _ = svc.PowerPointAddSlide(
                    sessionId,
                    payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."),
                    payload["body"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'body'."),
                    payload["bodyType"]?.GetValue<string>() ?? "text"),
            ["powerpoint_set_text_style"] = static (svc, sessionId, payload) =>
                _ = svc.PowerPointSetTextStyle(
                    sessionId,
                    payload["slideIndex"]?.GetValue<int>() ?? 1,
                    payload["slot"]?.GetValue<int>() ?? 0,
                    payload["fontName"]?.GetValue<string>() ?? "Calibri",
                    payload["fontSize"]?.GetValue<int>() ?? 24,
                    payload["bold"]?.GetValue<bool>() ?? false,
                    payload["italic"]?.GetValue<bool>() ?? false,
                    payload["colorHex"]?.GetValue<string>() ?? "000000"),
            ["powerpoint_insert_slide_at"] = static (svc, sessionId, payload) =>
                _ = svc.PowerPointInsertSlideAt(sessionId, payload["index"]?.GetValue<int>() ?? 0, payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."), payload["body"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'body'.")),
            ["powerpoint_set_slide_title"] = static (svc, sessionId, payload) =>
                _ = svc.PowerPointSetSlideTitle(sessionId, payload["slideIndex"]?.GetValue<int>() ?? 0, payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'.")),
            ["powerpoint_set_slide_body"] = static (svc, sessionId, payload) =>
                _ = svc.PowerPointSetSlideBody(
                    sessionId,
                    payload["slideIndex"]?.GetValue<int>() ?? 0,
                    payload["body"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'body'."),
                    payload["bodyType"]?.GetValue<string>() ?? "text"),
            ["powerpoint_set_slide_notes"] = static (svc, sessionId, payload) =>
                _ = svc.PowerPointSetSlideNotes(sessionId, payload["slideIndex"]?.GetValue<int>() ?? 0, payload["notes"]?.GetValue<string>() ?? string.Empty),
            ["powerpoint_get_slide_notes"] = static (svc, sessionId, payload) =>
                _ = svc.PowerPointGetSlideNotes(sessionId, payload["slideIndex"]?.GetValue<int>() ?? 0),
            ["powerpoint_reorder_slide"] = static (svc, sessionId, payload) =>
                _ = svc.PowerPointReorderSlide(sessionId, payload["fromIndex"]?.GetValue<int>() ?? 0, payload["toIndex"]?.GetValue<int>() ?? 0),
            ["powerpoint_delete_slide"] = static (svc, sessionId, payload) =>
                _ = svc.PowerPointDeleteSlide(sessionId, payload["slideIndex"]?.GetValue<int>() ?? 0),
            ["powerpoint_add_bullet_slide"] = static (svc, sessionId, payload) =>
                _ = svc.PowerPointAddBulletSlide(sessionId, payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."), payload["bulletLines"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'bulletLines'.")),
            ["apply_style_preset"] = static (svc, sessionId, payload) =>
                _ = svc.ApplyStylePreset(sessionId, payload["preset"]?.GetValue<string>() ?? "default"),
            ["list_style_presets"] = static (svc, sessionId, _) => _ = svc.ListStylePresets(sessionId),
            ["apply_text_preset"] = static (svc, sessionId, payload) =>
                _ = svc.ApplyTextPreset(
                    sessionId,
                    payload["preset"]?.GetValue<string>() ?? "default",
                    payload["targetIndex"]?.GetValue<int>() ?? 1)
        };
    }
}
