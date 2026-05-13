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
            ["word_add_table"] = static (svc, sessionId, payload) =>
                svc.WordAddTable(sessionId, payload["rows"]?.GetValue<int>() ?? 0, payload["columns"]?.GetValue<int>() ?? 0),
            ["word_insert_paragraph_at"] = static (svc, sessionId, payload) =>
                svc.WordInsertParagraphAt(sessionId, payload["index"]?.GetValue<int>() ?? 0, payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'.")),
            ["word_replace_text"] = static (svc, sessionId, payload) =>
                svc.WordReplaceText(sessionId, payload["find"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'find'."), payload["replace"]?.GetValue<string>() ?? string.Empty, payload["matchCase"]?.GetValue<bool>() ?? false),
            ["word_add_heading"] = static (svc, sessionId, payload) =>
                svc.WordAddHeading(sessionId, payload["level"]?.GetValue<int>() ?? 1, payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'.")),
            ["word_add_bulleted_list"] = static (svc, sessionId, payload) =>
                svc.WordAddBulletedList(sessionId, payload["lines"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'lines'.")),
            ["excel_set_cell_value"] = static (svc, sessionId, payload) =>
                svc.ExcelSetCellValue(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."), payload["cellReference"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'cellReference'."), payload["value"]?.GetValue<string>() ?? string.Empty),
            ["excel_set_range_values"] = static (svc, sessionId, payload) =>
                svc.ExcelSetRangeValues(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."), payload["startCell"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'startCell'."), payload["valuesJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'valuesJson'.")),
            ["excel_set_formula"] = static (svc, sessionId, payload) =>
                svc.ExcelSetFormula(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."), payload["cellReference"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'cellReference'."), payload["formula"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'formula'.")),
            ["excel_add_worksheet"] = static (svc, sessionId, payload) =>
                svc.ExcelAddWorksheet(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'.")),
            ["powerpoint_add_slide"] = static (svc, sessionId, payload) =>
                svc.PowerPointAddSlide(sessionId, payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."), payload["body"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'body'.")),
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
                svc.ApplyStylePreset(sessionId, payload["preset"]?.GetValue<string>() ?? "default")
        };
    }
}
