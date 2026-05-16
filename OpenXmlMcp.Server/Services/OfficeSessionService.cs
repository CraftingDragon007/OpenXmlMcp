using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlMcp.Server.Models;

namespace OpenXmlMcp.Server.Services;

/// <summary>
/// Thin coordinator that exposes the full public API surface consumed by <see cref="Tools.OfficeTools"/>
/// and <see cref="OfficeBatchDispatcher"/>. All document-type-specific logic is delegated to
/// <see cref="WordDocumentService"/>, <see cref="ExcelDocumentService"/>, and
/// <see cref="PowerPointDocumentService"/>. Cross-cutting concerns (session lifecycle, undo,
/// batch execution, style presets, validation) are handled here or via <see cref="SessionManager"/>.
/// </summary>
public class OfficeSessionService(
    SessionManager sessionManager,
    WordDocumentService wordService,
    ExcelDocumentService excelService,
    PowerPointDocumentService powerPointService)
{
    // -------------------------------------------------------------------------
    // Session lifecycle
    // -------------------------------------------------------------------------

    public string OpenDocument(string filePath, bool readOnly = false)
        => sessionManager.OpenDocument(filePath, readOnly);

    public string CreateDocument(string filePath, string documentType)
    {
        var normalizedPath = System.IO.Path.GetFullPath(filePath);
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(normalizedPath) ?? System.IO.Directory.GetCurrentDirectory());

        var type = ParseDocumentType(documentType);
        CreateEmptyDocument(normalizedPath, type);

        return sessionManager.OpenDocument(normalizedPath);
    }

    public string SaveDocument(string sessionId)
    {
        _ = sessionManager.GetSession(sessionId);
        return SessionManager.BuildMutationResult("save_document", new { sessionId });
    }

    public string CloseDocument(string sessionId)
        => sessionManager.CloseDocument(sessionId);

    public string GetDocumentInfo(string sessionId)
    {
        var session = sessionManager.GetSession(sessionId);
        var fileInfo = new System.IO.FileInfo(session.FilePath);

        var payload = new
        {
            session.Id,
            session.FilePath,
            documentType = session.DocumentType.ToString(),
            session.IsReadOnly,
            session.CreatedAtUtc,
            exists = fileInfo.Exists,
            sizeBytes = fileInfo.Exists ? fileInfo.Length : 0
        };

        return JsonSerializer.Serialize(payload);
    }

    public string GetOperationHistory(string sessionId)
        => sessionManager.GetOperationHistory(sessionId);

    public string UndoLastChange(string sessionId)
        => sessionManager.UndoLastChange(sessionId);

    // -------------------------------------------------------------------------
    // Cross-document operations
    // -------------------------------------------------------------------------

    public string ListStructure(string sessionId)
    {
        var session = sessionManager.GetSession(sessionId);

        object payload = session.DocumentType switch
        {
            OfficeDocumentType.Word => wordService.ListStructure(session.FilePath),
            OfficeDocumentType.Excel => excelService.ListStructure(session.FilePath),
            OfficeDocumentType.PowerPoint => powerPointService.ListStructure(session.FilePath),
            _ => throw new InvalidOperationException("Unsupported document type.")
        };

        return JsonSerializer.Serialize(payload);
    }

    public string FindText(string sessionId, string query)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(query);
        var session = sessionManager.GetSession(sessionId);

        object payload = session.DocumentType switch
        {
            OfficeDocumentType.Word => wordService.FindText(session.FilePath, query),
            OfficeDocumentType.Excel => excelService.FindText(session.FilePath, query),
            OfficeDocumentType.PowerPoint => powerPointService.FindText(session.FilePath, query),
            _ => throw new InvalidOperationException("Unsupported document type.")
        };

        return JsonSerializer.Serialize(payload);
    }

    public string ValidateOperation(string sessionId, string operationName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(operationName);
        var session = sessionManager.GetSession(sessionId);

        var normalized = operationName.Trim().ToLowerInvariant();
        var spec = OfficeOperationRegistry.TryGet(normalized);

        var warnings = new List<string>();
        var isValid = true;

        if (spec is null)
        {
            isValid = false;
            warnings.Add($"Unknown operation '{operationName}'.");
        }

        if (spec?.RequiresWrite == true && session.IsReadOnly)
        {
            isValid = false;
            warnings.Add("Session is read-only.");
        }

        if (spec?.ExpectedType is not null && session.DocumentType != spec.ExpectedType.Value)
        {
            isValid = false;
            warnings.Add($"Operation '{operationName}' expects '{spec.ExpectedType.Value}' document type.");
        }

        return JsonSerializer.Serialize(new { sessionId, operationName, isValid, warnings });
    }

    public string BatchExecute(string sessionId, string operationsJson)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(operationsJson);
        _ = sessionManager.GetSession(sessionId);

        var root = JsonNode.Parse(operationsJson) ?? throw new InvalidOperationException("Invalid operations JSON.");
        JsonArray operations;
        if (root is JsonArray directArray)
        {
            operations = directArray;
        }
        else if (root is JsonObject obj && obj["operations"] is JsonArray wrappedArray)
        {
            operations = wrappedArray;
        }
        else
        {
            throw new InvalidOperationException("Invalid batch payload. Expected a JSON array of operation objects or an object with an 'operations' array.");
        }

        var executed = 0;
        var failures = new List<object>();
        var results = new List<object>();

        for (var index = 0; index < operations.Count; index++)
        {
            var op = operations[index];
            if (op is null)
            {
                failures.Add(new { operation = string.Empty, index, errorCode = "InvalidBatchPayload", error = "Operation entry is null." });
                continue;
            }

            if (op is not JsonObject operationObject)
            {
                failures.Add(new { operation = string.Empty, index, errorCode = "InvalidBatchPayload", error = "Operation entry must be a JSON object." });
                continue;
            }

            var operation = operationObject["operation"]?.GetValue<string>()
                ?? operationObject["operationName"]?.GetValue<string>()
                ?? string.Empty;

            if (string.IsNullOrWhiteSpace(operation))
            {
                failures.Add(new { operation = string.Empty, index, errorCode = "MissingField", error = "Operation entry must include 'operation' (or legacy alias 'operationName')." });
                continue;
            }

            try
            {
                OfficeBatchDispatcher.Dispatch(this, sessionId, operation, operationObject);
                executed++;
                results.Add(new { operation, index, ok = true });
            }
            catch (Exception ex)
            {
                failures.Add(new { operation, index, errorCode = GetErrorCode(ex), error = ex.Message });
                results.Add(new { operation, index, ok = false, errorCode = GetErrorCode(ex), error = ex.Message });
            }
        }

        return JsonSerializer.Serialize(new
        {
            total = operations.Count,
            executed,
            failed = failures.Count,
            failures,
            results
        });
    }

    // -------------------------------------------------------------------------
    // Style presets (cross-suite)
    // -------------------------------------------------------------------------

    public string ApplyStylePreset(string sessionId, string preset = "default")
    {
        var session = sessionManager.GetSession(sessionId);
        var writableSession = sessionManager.GetWritableSession(sessionId, session.DocumentType);
        var normalizedPreset = string.IsNullOrWhiteSpace(preset) ? "default" : preset.Trim().ToLowerInvariant();
        ValidateStylePresetForDocumentType(session.DocumentType, normalizedPreset);

        sessionManager.ExecuteWriteOperation(writableSession, "apply_style_preset", () =>
        {
            switch (session.DocumentType)
            {
                case OfficeDocumentType.Word:
                    wordService.ApplyStylePreset(session.FilePath);
                    break;
                case OfficeDocumentType.Excel:
                    excelService.ApplyStylePreset(session.FilePath);
                    break;
                case OfficeDocumentType.PowerPoint:
                    powerPointService.ApplyStylePreset(session.FilePath, normalizedPreset);
                    break;
                default:
                    throw new InvalidOperationException("Unsupported document type.");
            }
        });

        return SessionManager.BuildMutationResult("apply_style_preset", new { preset = normalizedPreset });
    }

    public string ListStylePresets(string sessionId)
    {
        var session = sessionManager.GetSession(sessionId);
        var presets = session.DocumentType switch
        {
            OfficeDocumentType.Word => new[] { "default" },
            OfficeDocumentType.Excel => new[] { "default" },
            OfficeDocumentType.PowerPoint => PowerPointDocumentService.GetAllowedPresets(),
            _ => Array.Empty<string>()
        };

        return JsonSerializer.Serialize(new { documentType = session.DocumentType.ToString(), presets });
    }

    public string ApplyTextPreset(string sessionId, string preset, int targetIndex = 1)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(preset);
        var session = sessionManager.GetSession(sessionId);
        var normalized = preset.Trim().ToLowerInvariant();

        switch (session.DocumentType)
        {
            case OfficeDocumentType.Word:
                wordService.SetParagraphStyle(sessionId, targetIndex, BuildTextStyleFromPreset(normalized, isTitleSlot: true));
                break;
            case OfficeDocumentType.Excel:
                excelService.SetCellStyleInternal(sessionId, "Sheet1", "A1", BuildTextStyleFromPreset(normalized, isTitleSlot: true));
                break;
            case OfficeDocumentType.PowerPoint:
                powerPointService.SetTextStyleInternal(sessionId, targetIndex, 0, BuildTextStyleFromPreset(normalized, isTitleSlot: true));
                if (normalized is "title" or "subtitle")
                {
                    powerPointService.SetTextStyleInternal(sessionId, targetIndex, 1, BuildTextStyleFromPreset(normalized, isTitleSlot: false));
                }

                break;
            default:
                throw new InvalidOperationException("Unsupported document type.");
        }

        return SessionManager.BuildMutationResult("apply_text_preset", new { preset = normalized, targetIndex });
    }

    // -------------------------------------------------------------------------
    // Word operations (delegated)
    // -------------------------------------------------------------------------

    public string WordAppendParagraph(string sessionId, string text)
        => wordService.AppendParagraph(sessionId, text);

    public string WordSetParagraphStyle(string sessionId, int paragraphIndex, string fontName, int fontSize, bool bold, bool italic, string colorHex)
        => wordService.SetParagraphStyle(sessionId, paragraphIndex, fontName, fontSize, bold, italic, colorHex);

    public string WordSetParagraphSpacing(string sessionId, int paragraphIndex, int beforePt, int afterPt, double lineSpacing)
        => wordService.SetParagraphSpacing(sessionId, paragraphIndex, beforePt, afterPt, lineSpacing);

    public string WordSetDocumentSpacingPreset(string sessionId, string preset = "normal")
        => wordService.SetDocumentSpacingPreset(sessionId, preset);

    public string WordAddTable(string sessionId, int rows, int columns)
        => wordService.AddTable(sessionId, rows, columns);

    public string WordSetTableCell(string sessionId, int tableIndex, int rowIndex, int columnIndex, string text)
        => wordService.SetTableCell(sessionId, tableIndex, rowIndex, columnIndex, text);

    public string WordGetTableCell(string sessionId, int tableIndex, int rowIndex, int columnIndex)
        => wordService.GetTableCell(sessionId, tableIndex, rowIndex, columnIndex);

    public string WordGetParagraphInfo(string sessionId, int paragraphIndex)
        => wordService.GetParagraphInfo(sessionId, paragraphIndex);

    public string WordInsertParagraphAt(string sessionId, int index, string text)
        => wordService.InsertParagraphAt(sessionId, index, text);

    public string WordReplaceText(string sessionId, string find, string replace, bool matchCase = false)
        => wordService.ReplaceText(sessionId, find, replace, matchCase);

    public string WordAddHeading(string sessionId, int level, string text)
        => wordService.AddHeading(sessionId, level, text);

    public string WordAddBulletedList(string sessionId, string lines, string bulletStyle = "disc")
        => wordService.AddBulletedList(sessionId, lines, bulletStyle);

    public string WordAddNumberedList(string sessionId, string lines, string numberStyle = "decimal-dot")
        => wordService.AddNumberedList(sessionId, lines, numberStyle);

    public string WordAddStructuredList(string sessionId, string itemsJson)
        => wordService.AddStructuredList(sessionId, itemsJson);

    public string WordInsertParagraphAfterText(string sessionId, string anchorText, string text, int occurrence = 1, bool matchCase = false)
        => wordService.InsertParagraphAfterText(sessionId, anchorText, text, occurrence, matchCase);

    public string WordInsertTextAfterText(string sessionId, string anchorText, string text, int occurrence = 1, bool matchCase = false)
        => wordService.InsertTextAfterText(sessionId, anchorText, text, occurrence, matchCase);

    public string WordListStyles(string sessionId)
        => wordService.ListStyles(sessionId);

    public string WordApplyStyleByName(string sessionId, int paragraphIndex, string styleName)
        => wordService.ApplyStyleByName(sessionId, paragraphIndex, styleName);

    public string WordCreateOrUpdateStyle(string sessionId, string styleName, string styleJson)
        => wordService.CreateOrUpdateStyle(sessionId, styleName, styleJson);

    public string WordApplyCharacterStyleToText(string sessionId, string anchorText, string styleName, int occurrence = 1, bool matchCase = false)
        => wordService.ApplyCharacterStyleToText(sessionId, anchorText, styleName, occurrence, matchCase);

    public string WordApplyCharacterStyleToAll(string sessionId, string queriesJson, string styleName, bool matchCase = false, bool wholeWord = true)
        => wordService.ApplyCharacterStyleToAll(sessionId, queriesJson, styleName, matchCase, wholeWord);

    public string WordApplyCharacterStyleByPattern(string sessionId, string pattern, string styleName, bool matchCase = true, int maxMatches = 5000)
        => wordService.ApplyCharacterStyleByPattern(sessionId, pattern, styleName, matchCase, maxMatches);

    public string WordInsertAfterHeading(string sessionId, string headingText, string text, int occurrence = 1, bool matchCase = false)
        => wordService.InsertAfterHeading(sessionId, headingText, text, occurrence, matchCase);

    public string WordReplaceSection(string sessionId, string headingText, string replacementJson, int occurrence = 1, bool matchCase = false)
        => wordService.ReplaceSection(sessionId, headingText, replacementJson, occurrence, matchCase);

    public string WordInsertTableOfContents(string sessionId, int paragraphIndex, int minLevel = 1, int maxLevel = 3)
        => wordService.InsertTableOfContents(sessionId, paragraphIndex, minLevel, maxLevel);

    public string WordInsertPageBreakAfter(string sessionId, int paragraphIndex)
        => wordService.InsertPageBreakAfter(sessionId, paragraphIndex);

    public string WordSetHeader(string sessionId, string text, int sectionIndex = 1)
        => wordService.SetHeader(sessionId, text, sectionIndex);

    public string WordSetFooter(string sessionId, string text, int sectionIndex = 1)
        => wordService.SetFooter(sessionId, text, sectionIndex);

    public string WordApplyTableStyle(string sessionId, int tableIndex, string styleName)
        => wordService.ApplyTableStyle(sessionId, tableIndex, styleName);

    public string WordFormatTableHeaderRow(string sessionId, int tableIndex, bool bold = true, string? shadingFill = null, string? fontName = null, string? colorHex = null)
        => wordService.FormatTableHeaderRow(sessionId, tableIndex, bold, shadingFill, fontName, colorHex);

    public string WordSetTableValues(string sessionId, int tableIndex, string valuesJson, int startRow = 1, int startColumn = 1)
        => wordService.SetTableValues(sessionId, tableIndex, valuesJson, startRow, startColumn);

    public string WordListParagraphRuns(string sessionId, int paragraphIndex)
        => wordService.ListParagraphRuns(sessionId, paragraphIndex);

    // -------------------------------------------------------------------------
    // Excel operations (delegated)
    // -------------------------------------------------------------------------

    public string ExcelSetCellValue(string sessionId, string sheetName, string cellReference, string value)
        => excelService.SetCellValue(sessionId, sheetName, cellReference, value);

    public string ExcelSetCellStyle(string sessionId, string sheetName, string cellReference, string fontName, int fontSize, bool bold, bool italic, string colorHex)
        => excelService.SetCellStyle(sessionId, sheetName, cellReference, fontName, fontSize, bold, italic, colorHex);

    public string ExcelGetCellValue(string sessionId, string sheetName, string cellReference)
        => excelService.GetCellValue(sessionId, sheetName, cellReference);

    public string ExcelGetCellInfo(string sessionId, string sheetName, string cellReference)
        => excelService.GetCellInfo(sessionId, sheetName, cellReference);

    public string ExcelGetCellStyle(string sessionId, string sheetName, string cellReference)
        => excelService.GetCellStyle(sessionId, sheetName, cellReference);

    public string ExcelGetUsedRange(string sessionId, string sheetName)
        => excelService.GetUsedRange(sessionId, sheetName);

    public string ExcelSetRangeValues(string sessionId, string sheetName, string startCell, string valuesJson)
        => excelService.SetRangeValues(sessionId, sheetName, startCell, valuesJson);

    public string ExcelSetFormula(string sessionId, string sheetName, string cellReference, string formula)
        => excelService.SetFormula(sessionId, sheetName, cellReference, formula);

    public string ExcelGetFormula(string sessionId, string sheetName, string cellReference)
        => excelService.GetFormula(sessionId, sheetName, cellReference);

    public string ExcelAddWorksheet(string sessionId, string sheetName)
        => excelService.AddWorksheet(sessionId, sheetName);

    // -------------------------------------------------------------------------
    // PowerPoint operations (delegated)
    // -------------------------------------------------------------------------

    public string PowerPointAddSlide(string sessionId, string title, string body, string bodyType = "text")
        => powerPointService.AddSlide(sessionId, title, body, bodyType);

    public string PowerPointAddBulletSlide(string sessionId, string title, string bulletLines)
        => powerPointService.AddBulletSlide(sessionId, title, bulletLines);

    public string PowerPointInsertSlideAt(string sessionId, int index, string title, string body)
        => powerPointService.InsertSlideAt(sessionId, index, title, body);

    public string PowerPointSetSlideTitle(string sessionId, int slideIndex, string title)
        => powerPointService.SetSlideTitle(sessionId, slideIndex, title);

    public string PowerPointSetSlideBody(string sessionId, int slideIndex, string body, string bodyType = "text")
        => powerPointService.SetSlideBody(sessionId, slideIndex, body, bodyType);

    public string PowerPointSetSlideNotes(string sessionId, int slideIndex, string notes)
        => powerPointService.SetSlideNotes(sessionId, slideIndex, notes);

    public string PowerPointGetSlideNotes(string sessionId, int slideIndex)
        => powerPointService.GetSlideNotes(sessionId, slideIndex);

    public string PowerPointSetTextStyle(string sessionId, int slideIndex, int slot, string fontName, int fontSize, bool bold, bool italic, string colorHex)
        => powerPointService.SetTextStyle(sessionId, slideIndex, slot, fontName, fontSize, bold, italic, colorHex);

    public string PowerPointGetTextStyle(string sessionId, int slideIndex, int slot)
        => powerPointService.GetTextStyle(sessionId, slideIndex, slot);

    public string PowerPointReorderSlide(string sessionId, int fromIndex, int toIndex)
        => powerPointService.ReorderSlide(sessionId, fromIndex, toIndex);

    public string PowerPointDeleteSlide(string sessionId, int slideIndex)
        => powerPointService.DeleteSlide(sessionId, slideIndex);

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    private void CreateEmptyDocument(string filePath, OfficeDocumentType documentType)
    {
        if (System.IO.File.Exists(filePath))
        {
            System.IO.File.Delete(filePath);
        }

        switch (documentType)
        {
            case OfficeDocumentType.Word:
                wordService.InitializeEmptyDocument(filePath);
                break;
            case OfficeDocumentType.Excel:
                excelService.InitializeEmptyDocument(filePath);
                break;
            case OfficeDocumentType.PowerPoint:
                powerPointService.InitializeEmptyDocument(filePath);
                break;
            default:
                throw new InvalidOperationException("Unsupported document type.");
        }
    }

    private static void ValidateStylePresetForDocumentType(OfficeDocumentType documentType, string preset)
    {
        var allowed = documentType switch
        {
            OfficeDocumentType.Word => new[] { "default" },
            OfficeDocumentType.Excel => new[] { "default" },
            OfficeDocumentType.PowerPoint => PowerPointDocumentService.GetAllowedPresets(),
            _ => Array.Empty<string>()
        };

        if (!allowed.Contains(preset, StringComparer.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"Invalid preset '{preset}' for {documentType}. Allowed presets: {string.Join(", ", allowed)}.");
        }
    }

    private static TextStyle BuildTextStyleFromPreset(string preset, bool isTitleSlot)
    {
        return preset switch
        {
            "title" => isTitleSlot
                ? new TextStyle("Calibri", 36, true, false, "1F497D")
                : new TextStyle("Calibri", 22, false, false, "4F4F4F"),
            "subtitle" => isTitleSlot
                ? new TextStyle("Calibri", 28, true, false, "1F497D")
                : new TextStyle("Calibri", 20, false, true, "666666"),
            _ => new TextStyle("Calibri", 18, false, false, "000000")
        };
    }

    private static OfficeDocumentType ParseDocumentType(string documentType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(documentType);
        return documentType.Trim().ToLowerInvariant() switch
        {
            "docx" or "word" => OfficeDocumentType.Word,
            "xlsx" or "excel" => OfficeDocumentType.Excel,
            "pptx" or "powerpoint" or "ppt" => OfficeDocumentType.PowerPoint,
            _ => throw new InvalidOperationException($"Unsupported document type '{documentType}'.")
        };
    }

    private static string GetErrorCode(Exception ex)
    {
        if (ex.Message.Contains("Invalid batch payload", StringComparison.OrdinalIgnoreCase)
            || ex.Message.Contains("Operation entry", StringComparison.OrdinalIgnoreCase))
        {
            return "InvalidBatchPayload";
        }

        if (ex.Message.Contains("Missing '", StringComparison.OrdinalIgnoreCase))
        {
            return "MissingField";
        }

        if (ex.Message.Contains("not supported in batch execution", StringComparison.OrdinalIgnoreCase)
            || ex.Message.Contains("Unsupported batch operation", StringComparison.OrdinalIgnoreCase))
        {
            return "UnknownOperation";
        }

        if (ex.Message.Contains("read-only", StringComparison.OrdinalIgnoreCase))
        {
            return "ReadOnlySession";
        }

        if (ex.Message.Contains("expects", StringComparison.OrdinalIgnoreCase)
            || ex.Message.Contains("document type", StringComparison.OrdinalIgnoreCase))
        {
            return "WrongDocumentType";
        }

        if (ex.Message.Contains("Anchor text", StringComparison.OrdinalIgnoreCase))
        {
            return "AnchorNotFound";
        }

        if (ex.Message.Contains("Invalid preset", StringComparison.OrdinalIgnoreCase))
        {
            return "InvalidPreset";
        }

        return ex switch
        {
            InvalidOperationException => "InvalidOperation",
            ArgumentException => "InvalidArgument",
            System.Text.Json.JsonException => "InvalidJson",
            _ => "UnhandledError"
        };
    }
}
