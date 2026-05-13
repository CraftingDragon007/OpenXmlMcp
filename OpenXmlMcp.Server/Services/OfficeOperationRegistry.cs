using OpenXmlMcp.Server.Models;

namespace OpenXmlMcp.Server.Services;

internal sealed record OfficeOperationSpec(string CanonicalName, OfficeDocumentType? ExpectedType, bool RequiresWrite, params string[] Aliases);

internal static class OfficeOperationRegistry
{
    private static readonly IReadOnlyDictionary<string, OfficeOperationSpec> Registry = Build();

    public static OfficeOperationSpec? TryGet(string normalizedOperation)
    {
        return Registry.TryGetValue(normalizedOperation, out var spec) ? spec : null;
    }

    public static IReadOnlyDictionary<string, OfficeOperationSpec> Build()
    {
        var specs = new List<OfficeOperationSpec>
        {
            new("open_document", null, false),
            new("create_document", null, false),
            new("save_document", null, false),
            new("close_document", null, false),
            new("get_document_info", null, false),
            new("list_structure", null, false),
            new("find_text", null, false),
            new("validate_operation", null, false),
            new("get_operation_history", null, false),
            new("undo_last_change", null, true),
            new("batch_execute", null, true),
            new("apply_style_preset", null, true),
            new("list_style_presets", null, false),
            new("apply_text_preset", null, true),
            new("word_append_paragraph", OfficeDocumentType.Word, true),
            new("word_set_paragraph_style", OfficeDocumentType.Word, true),
            new("word_insert_paragraph_at", OfficeDocumentType.Word, true),
            new("word_replace_text", OfficeDocumentType.Word, true),
            new("word_add_heading", OfficeDocumentType.Word, true),
            new("word_add_bulleted_list", OfficeDocumentType.Word, true),
            new("word_add_numbered_list", OfficeDocumentType.Word, true),
            new("word_add_structured_list", OfficeDocumentType.Word, true),
            new("word_set_paragraph_spacing", OfficeDocumentType.Word, true),
            new("word_set_document_spacing_preset", OfficeDocumentType.Word, true),
            new("word_insert_paragraph_after_text", OfficeDocumentType.Word, true),
            new("word_insert_text_after_text", OfficeDocumentType.Word, true),
            new("word_list_styles", OfficeDocumentType.Word, false),
            new("word_apply_style_by_name", OfficeDocumentType.Word, true),
            new("word_create_or_update_style", OfficeDocumentType.Word, true),
            new("word_add_table", OfficeDocumentType.Word, true),
            new("word_set_table_cell", OfficeDocumentType.Word, true),
            new("word_get_table_cell", OfficeDocumentType.Word, false),
            new("excel_set_cell_value", OfficeDocumentType.Excel, true),
            new("excel_set_cell_style", OfficeDocumentType.Excel, true),
            new("excel_get_cell_value", OfficeDocumentType.Excel, false),
            new("excel_get_used_range", OfficeDocumentType.Excel, false),
            new("excel_set_range_values", OfficeDocumentType.Excel, true),
            new("excel_set_formula", OfficeDocumentType.Excel, true),
            new("excel_get_formula", OfficeDocumentType.Excel, false),
            new("excel_add_worksheet", OfficeDocumentType.Excel, true),
            new("powerpoint_add_slide", OfficeDocumentType.PowerPoint, true, "power_point_add_slide"),
            new("powerpoint_set_text_style", OfficeDocumentType.PowerPoint, true, "power_point_set_text_style"),
            new("powerpoint_insert_slide_at", OfficeDocumentType.PowerPoint, true, "power_point_insert_slide_at"),
            new("powerpoint_set_slide_title", OfficeDocumentType.PowerPoint, true, "power_point_set_slide_title"),
            new("powerpoint_set_slide_body", OfficeDocumentType.PowerPoint, true, "power_point_set_slide_body"),
            new("powerpoint_set_slide_notes", OfficeDocumentType.PowerPoint, true, "power_point_set_slide_notes"),
            new("powerpoint_get_slide_notes", OfficeDocumentType.PowerPoint, false, "power_point_get_slide_notes"),
            new("powerpoint_reorder_slide", OfficeDocumentType.PowerPoint, true, "power_point_reorder_slide"),
            new("powerpoint_delete_slide", OfficeDocumentType.PowerPoint, true, "power_point_delete_slide"),
            new("powerpoint_add_bullet_slide", OfficeDocumentType.PowerPoint, true, "power_point_add_bullet_slide")
        };

        var map = new Dictionary<string, OfficeOperationSpec>(StringComparer.OrdinalIgnoreCase);
        foreach (var spec in specs)
        {
            map[spec.CanonicalName] = spec;
            foreach (var alias in spec.Aliases)
            {
                map[alias] = spec;
            }
        }

        return map;
    }
}
