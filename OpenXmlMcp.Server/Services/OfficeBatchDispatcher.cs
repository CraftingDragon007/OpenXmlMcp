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
                svc.WordAppendParagraph(sessionId, payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'.")),
            ["word_set_paragraph_style"] = static (svc, sessionId, payload) =>
                svc.WordSetParagraphStyle(
                    sessionId,
                    payload["paragraphIndex"]?.GetValue<int>() ?? 1,
                    payload["fontName"]?.GetValue<string>() ?? "Calibri",
                    payload["fontSize"]?.GetValue<int>() ?? 18,
                    payload["bold"]?.GetValue<bool>() ?? false,
                    payload["italic"]?.GetValue<bool>() ?? false,
                    payload["colorHex"]?.GetValue<string>() ?? "000000"),
            ["word_add_table"] = static (svc, sessionId, payload) =>
                svc.WordAddTable(sessionId, payload["rows"]?.GetValue<int>() ?? 0, payload["columns"]?.GetValue<int>() ?? 0),
            ["word_insert_paragraph_at"] = static (svc, sessionId, payload) =>
                svc.WordInsertParagraphAt(sessionId, payload["index"]?.GetValue<int>() ?? 0, payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'.")),
            ["word_replace_text"] = static (svc, sessionId, payload) =>
                svc.WordReplaceText(sessionId, payload["find"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'find'."), payload["replace"]?.GetValue<string>() ?? string.Empty, payload["matchCase"]?.GetValue<bool>() ?? false),
            ["word_add_heading"] = static (svc, sessionId, payload) =>
                svc.WordAddHeading(sessionId, payload["level"]?.GetValue<int>() ?? 1, payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'.")),
            ["word_add_bulleted_list"] = static (svc, sessionId, payload) =>
                svc.WordAddBulletedList(
                    sessionId,
                    payload["lines"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'lines'."),
                    payload["bulletStyle"]?.GetValue<string>() ?? "disc"),
            ["word_add_numbered_list"] = static (svc, sessionId, payload) =>
                svc.WordAddNumberedList(
                    sessionId,
                    payload["lines"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'lines'."),
                    payload["numberStyle"]?.GetValue<string>() ?? "decimal-dot"),
            ["word_add_structured_list"] = static (svc, sessionId, payload) =>
                svc.WordAddStructuredList(
                    sessionId,
                    payload["itemsJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'itemsJson'.")),
            ["word_set_paragraph_spacing"] = static (svc, sessionId, payload) =>
                svc.WordSetParagraphSpacing(
                    sessionId,
                    payload["paragraphIndex"]?.GetValue<int>() ?? 1,
                    payload["beforePt"]?.GetValue<int>() ?? 0,
                    payload["afterPt"]?.GetValue<int>() ?? 8,
                    payload["lineSpacing"]?.GetValue<double>() ?? 1.15),
            ["word_set_document_spacing_preset"] = static (svc, sessionId, payload) =>
                svc.WordSetDocumentSpacingPreset(
                    sessionId,
                    payload["preset"]?.GetValue<string>() ?? "normal"),
            ["word_insert_paragraph_after_text"] = static (svc, sessionId, payload) =>
                svc.WordInsertParagraphAfterText(
                    sessionId,
                    payload["anchorText"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'anchorText'."),
                    payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'."),
                    payload["occurrence"]?.GetValue<int>() ?? 1,
                    payload["matchCase"]?.GetValue<bool>() ?? false),
            ["word_insert_text_after_text"] = static (svc, sessionId, payload) =>
                svc.WordInsertTextAfterText(
                    sessionId,
                    payload["anchorText"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'anchorText'."),
                    payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'."),
                    payload["occurrence"]?.GetValue<int>() ?? 1,
                    payload["matchCase"]?.GetValue<bool>() ?? false),
            ["word_apply_style_by_name"] = static (svc, sessionId, payload) =>
                svc.WordApplyStyleByName(
                    sessionId,
                    payload["paragraphIndex"]?.GetValue<int>() ?? 1,
                    payload["styleName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'styleName'.")),
            ["word_create_or_update_style"] = static (svc, sessionId, payload) =>
                svc.WordCreateOrUpdateStyle(
                    sessionId,
                    payload["styleName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'styleName'."),
                    payload["styleJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'styleJson'.")),
            ["word_list_styles"] = static (svc, sessionId, _) => _ = svc.WordListStyles(sessionId),
            ["excel_set_cell_value"] = static (svc, sessionId, payload) =>
                svc.ExcelSetCellValue(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."), payload["cellReference"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'cellReference'."), payload["value"]?.GetValue<string>() ?? string.Empty),
            ["excel_set_cell_style"] = static (svc, sessionId, payload) =>
                svc.ExcelSetCellStyle(
                    sessionId,
                    payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."),
                    payload["cellReference"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'cellReference'."),
                    payload["fontName"]?.GetValue<string>() ?? "Calibri",
                    payload["fontSize"]?.GetValue<int>() ?? 11,
                    payload["bold"]?.GetValue<bool>() ?? false,
                    payload["italic"]?.GetValue<bool>() ?? false,
                    payload["colorHex"]?.GetValue<string>() ?? "000000"),
            ["excel_set_range_values"] = static (svc, sessionId, payload) =>
                svc.ExcelSetRangeValues(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."), payload["startCell"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'startCell'."), payload["valuesJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'valuesJson'.")),
            ["excel_set_formula"] = static (svc, sessionId, payload) =>
                svc.ExcelSetFormula(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."), payload["cellReference"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'cellReference'."), payload["formula"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'formula'.")),
            ["excel_add_worksheet"] = static (svc, sessionId, payload) =>
                svc.ExcelAddWorksheet(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'.")),
            ["powerpoint_add_slide"] = static (svc, sessionId, payload) =>
                svc.PowerPointAddSlide(sessionId, payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."), payload["body"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'body'.")),
            ["powerpoint_set_text_style"] = static (svc, sessionId, payload) =>
                svc.PowerPointSetTextStyle(
                    sessionId,
                    payload["slideIndex"]?.GetValue<int>() ?? 1,
                    payload["slot"]?.GetValue<int>() ?? 0,
                    payload["fontName"]?.GetValue<string>() ?? "Calibri",
                    payload["fontSize"]?.GetValue<int>() ?? 24,
                    payload["bold"]?.GetValue<bool>() ?? false,
                    payload["italic"]?.GetValue<bool>() ?? false,
                    payload["colorHex"]?.GetValue<string>() ?? "000000"),
            ["powerpoint_insert_slide_at"] = static (svc, sessionId, payload) =>
                svc.PowerPointInsertSlideAt(sessionId, payload["index"]?.GetValue<int>() ?? 0, payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."), payload["body"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'body'.")),
            ["powerpoint_set_slide_title"] = static (svc, sessionId, payload) =>
                svc.PowerPointSetSlideTitle(sessionId, payload["slideIndex"]?.GetValue<int>() ?? 0, payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'.")),
            ["powerpoint_set_slide_body"] = static (svc, sessionId, payload) =>
                svc.PowerPointSetSlideBody(sessionId, payload["slideIndex"]?.GetValue<int>() ?? 0, payload["body"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'body'.")),
            ["powerpoint_reorder_slide"] = static (svc, sessionId, payload) =>
                svc.PowerPointReorderSlide(sessionId, payload["fromIndex"]?.GetValue<int>() ?? 0, payload["toIndex"]?.GetValue<int>() ?? 0),
            ["powerpoint_delete_slide"] = static (svc, sessionId, payload) =>
                svc.PowerPointDeleteSlide(sessionId, payload["slideIndex"]?.GetValue<int>() ?? 0),
            ["powerpoint_add_bullet_slide"] = static (svc, sessionId, payload) =>
                svc.PowerPointAddBulletSlide(sessionId, payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."), payload["bulletLines"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'bulletLines'.")),
            ["apply_style_preset"] = static (svc, sessionId, payload) =>
                svc.ApplyStylePreset(sessionId, payload["preset"]?.GetValue<string>() ?? "default"),
            ["list_style_presets"] = static (svc, sessionId, _) => _ = svc.ListStylePresets(sessionId),
            ["apply_text_preset"] = static (svc, sessionId, payload) =>
                svc.ApplyTextPreset(
                    sessionId,
                    payload["preset"]?.GetValue<string>() ?? "default",
                    payload["targetIndex"]?.GetValue<int>() ?? 1)
        };
    }
}
