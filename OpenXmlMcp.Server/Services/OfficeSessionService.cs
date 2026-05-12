using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlMcp.Server.Models;

namespace OpenXmlMcp.Server.Services;

public class OfficeSessionService
{
    private readonly ConcurrentDictionary<string, OfficeSession> _sessions = new();
    private readonly ConcurrentDictionary<string, Stack<string>> _sessionSnapshots = new();
    private readonly ConcurrentDictionary<string, List<OperationLogEntry>> _sessionOperationLog = new();
    private const long MaxOpenFileSizeBytes = 20 * 1024 * 1024;

    public string OpenDocument(string filePath, bool readOnly = false)
    {
        var normalizedPath = NormalizePath(filePath);

        if (!File.Exists(normalizedPath))
        {
            throw new FileNotFoundException("Office file not found.", normalizedPath);
        }

        var fileInfo = new FileInfo(normalizedPath);
        if (fileInfo.Length > MaxOpenFileSizeBytes)
        {
            throw new InvalidOperationException($"File size {fileInfo.Length} exceeds safety limit {MaxOpenFileSizeBytes} bytes.");
        }

        var session = new OfficeSession
        {
            Id = Guid.NewGuid().ToString("N"),
            FilePath = normalizedPath,
            DocumentType = GetDocumentTypeFromPath(normalizedPath),
            IsReadOnly = readOnly,
            CreatedAtUtc = DateTimeOffset.UtcNow
        };

        _sessions[session.Id] = session;
        _sessionSnapshots.TryAdd(session.Id, new Stack<string>());
        _sessionOperationLog.TryAdd(session.Id, new List<OperationLogEntry>());
        AppendOperationLog(session.Id, "open_document", $"Opened '{normalizedPath}' (readOnly={readOnly}).");
        return session.Id;
    }

    public string CreateDocument(string filePath, string documentType)
    {
        var normalizedPath = NormalizePath(filePath);
        Directory.CreateDirectory(Path.GetDirectoryName(normalizedPath) ?? Directory.GetCurrentDirectory());

        var type = ParseDocumentType(documentType);
        CreateEmptyDocument(normalizedPath, type);

        return OpenDocument(normalizedPath);
    }

    public void SaveDocument(string sessionId)
    {
        _ = GetSession(sessionId);
    }

    public void CloseDocument(string sessionId)
    {
        if (_sessionSnapshots.TryRemove(sessionId, out var snapshots))
        {
            foreach (var snapshotFile in snapshots)
            {
                DeleteIfExists(snapshotFile);
            }
        }

        _sessionOperationLog.TryRemove(sessionId, out _);

        if (!_sessions.TryRemove(sessionId, out _))
        {
            throw new InvalidOperationException($"Session '{sessionId}' was not found.");
        }
    }

    public string GetOperationHistory(string sessionId)
    {
        _ = GetSession(sessionId);
        var history = _sessionOperationLog.TryGetValue(sessionId, out var entries)
            ? entries.Select(x => new { x.TimestampUtc, x.OperationName, x.Message }).ToArray()
            : [];

        return JsonSerializer.Serialize(new { sessionId, count = history.Length, history });
    }

    public void UndoLastChange(string sessionId)
    {
        var session = GetSession(sessionId);
        if (session.IsReadOnly)
        {
            throw new InvalidOperationException("Session is read-only.");
        }

        if (!_sessionSnapshots.TryGetValue(sessionId, out var snapshots) || snapshots.Count == 0)
        {
            throw new InvalidOperationException("No checkpoint is available to undo.");
        }

        var snapshotPath = snapshots.Pop();
        File.Copy(snapshotPath, session.FilePath, overwrite: true);
        DeleteIfExists(snapshotPath);
        AppendOperationLog(sessionId, "undo_last_change", "Restored document from latest checkpoint.");
    }

    public string GetDocumentInfo(string sessionId)
    {
        var session = GetSession(sessionId);
        var fileInfo = new FileInfo(session.FilePath);

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

    public string ListStructure(string sessionId)
    {
        var session = GetSession(sessionId);

        object payload = session.DocumentType switch
        {
            OfficeDocumentType.Word => ListWordStructure(session.FilePath),
            OfficeDocumentType.Excel => ListExcelStructure(session.FilePath),
            OfficeDocumentType.PowerPoint => ListPowerPointStructure(session.FilePath),
            _ => throw new InvalidOperationException("Unsupported document type.")
        };

        return JsonSerializer.Serialize(payload);
    }

    public string FindText(string sessionId, string query)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(query);
        var session = GetSession(sessionId);

        object payload = session.DocumentType switch
        {
            OfficeDocumentType.Word => FindTextInWord(session.FilePath, query),
            OfficeDocumentType.Excel => FindTextInExcel(session.FilePath, query),
            OfficeDocumentType.PowerPoint => FindTextInPowerPoint(session.FilePath, query),
            _ => throw new InvalidOperationException("Unsupported document type.")
        };

        return JsonSerializer.Serialize(payload);
    }

    public string ValidateOperation(string sessionId, string operationName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(operationName);
        var session = GetSession(sessionId);

        var normalized = operationName.Trim().ToLowerInvariant();
        var requiresWrite = normalized is
            "word_append_paragraph" or
            "word_insert_paragraph_at" or
            "word_replace_text" or
            "word_add_heading" or
            "word_add_bulleted_list" or
            "word_add_table" or
            "excel_set_cell_value" or
            "excel_set_range_values" or
            "excel_set_formula" or
            "excel_add_worksheet" or
            "powerpoint_add_slide" or
            "power_point_add_slide" or
            "powerpoint_insert_slide_at" or
            "power_point_insert_slide_at" or
            "powerpoint_set_slide_title" or
            "power_point_set_slide_title" or
            "powerpoint_set_slide_body" or
            "power_point_set_slide_body" or
            "powerpoint_reorder_slide" or
            "power_point_reorder_slide" or
            "powerpoint_delete_slide" or
            "power_point_delete_slide" or
            "powerpoint_add_bullet_slide" or
            "power_point_add_bullet_slide" or
            "batch_execute" or
            "undo_last_change";
        var expectedType = normalized switch
        {
            "word_append_paragraph" => OfficeDocumentType.Word,
            "word_insert_paragraph_at" => OfficeDocumentType.Word,
            "word_replace_text" => OfficeDocumentType.Word,
            "word_add_heading" => OfficeDocumentType.Word,
            "word_add_bulleted_list" => OfficeDocumentType.Word,
            "word_add_table" => OfficeDocumentType.Word,
            "excel_set_cell_value" or "excel_get_cell_value" => OfficeDocumentType.Excel,
            "excel_set_range_values" => OfficeDocumentType.Excel,
            "excel_set_formula" or "excel_get_formula" or "excel_get_used_range" => OfficeDocumentType.Excel,
            "excel_add_worksheet" => OfficeDocumentType.Excel,
            "powerpoint_add_slide" or "power_point_add_slide" => OfficeDocumentType.PowerPoint,
            "powerpoint_insert_slide_at" or "power_point_insert_slide_at" => OfficeDocumentType.PowerPoint,
            "powerpoint_set_slide_title" or "power_point_set_slide_title" => OfficeDocumentType.PowerPoint,
            "powerpoint_set_slide_body" or "power_point_set_slide_body" => OfficeDocumentType.PowerPoint,
            "powerpoint_reorder_slide" or "power_point_reorder_slide" => OfficeDocumentType.PowerPoint,
            "powerpoint_delete_slide" or "power_point_delete_slide" => OfficeDocumentType.PowerPoint,
            "powerpoint_add_bullet_slide" or "power_point_add_bullet_slide" => OfficeDocumentType.PowerPoint,
            _ => (OfficeDocumentType?)null
        };

        var warnings = new List<string>();
        var isValid = true;

        if (expectedType is null && normalized is not "find_text" and not "list_structure" and not "get_document_info" and not "save_document" and not "close_document" and not "batch_execute" and not "get_operation_history" and not "undo_last_change")
        {
            isValid = false;
            warnings.Add($"Unknown operation '{operationName}'.");
        }

        if (requiresWrite && session.IsReadOnly)
        {
            isValid = false;
            warnings.Add("Session is read-only.");
        }

        if (expectedType is not null && session.DocumentType != expectedType.Value)
        {
            isValid = false;
            warnings.Add($"Operation '{operationName}' expects '{expectedType.Value}' document type.");
        }

        var payload = new
        {
            sessionId,
            operationName,
            isValid,
            warnings
        };

        return JsonSerializer.Serialize(payload);
    }

    public void WordAddTable(string sessionId, int rows, int columns)
    {
        if (rows <= 0 || columns <= 0)
        {
            throw new InvalidOperationException("Rows and columns must be greater than zero.");
        }

        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_add_table", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");

            var table = new W.Table();
            for (var r = 0; r < rows; r++)
            {
                var tableRow = new W.TableRow();
                for (var c = 0; c < columns; c++)
                {
                    tableRow.Append(new W.TableCell(new W.Paragraph(new W.Run(new W.Text(string.Empty)))));
                }

                table.Append(tableRow);
            }

            body.Append(table);
            document.MainDocumentPart.Document?.Save();
        });
    }

    public void WordInsertParagraphAt(string sessionId, int index, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);

        ExecuteWriteOperation(session, "word_insert_paragraph_at", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();

            if (index < 1 || index > paragraphs.Count + 1)
            {
                throw new InvalidOperationException($"Index {index} is out of range. Valid range is 1..{paragraphs.Count + 1}.");
            }

            var newParagraph = new W.Paragraph(new W.Run(new W.Text(text)));
            if (paragraphs.Count == 0 || index == paragraphs.Count + 1)
            {
                body.Append(newParagraph);
            }
            else
            {
                paragraphs[index - 1].InsertBeforeSelf(newParagraph);
            }

            document.MainDocumentPart.Document?.Save();
        });
    }

    public int WordReplaceText(string sessionId, string find, string replace, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(find);
        replace ??= string.Empty;
        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        var replacements = 0;

        ExecuteWriteOperation(session, "word_replace_text", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var textNodes = document.MainDocumentPart?.Document?.Descendants<W.Text>()
                ?? throw new InvalidOperationException("Word document text nodes are missing.");

            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            foreach (var textNode in textNodes)
            {
                if (string.IsNullOrEmpty(textNode.Text))
                {
                    continue;
                }

                var count = CountOccurrences(textNode.Text, find, comparison);
                if (count == 0)
                {
                    continue;
                }

                textNode.Text = ReplaceWithComparison(textNode.Text, find, replace, comparison);
                replacements += count;
            }

            document.MainDocumentPart.Document?.Save();
        });

        return replacements;
    }

    public void WordAddHeading(string sessionId, int level, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        if (level < 1 || level > 6)
        {
            throw new InvalidOperationException("Heading level must be between 1 and 6.");
        }

        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_add_heading", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");

            var heading = new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = $"Heading{level}" }),
                new W.Run(new W.Text(text)));

            body.Append(heading);
            document.MainDocumentPart.Document?.Save();
        });
    }

    public void WordAddBulletedList(string sessionId, string lines)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(lines);
        var items = lines.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (items.Length == 0)
        {
            throw new InvalidOperationException("At least one bullet line is required.");
        }

        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_add_bulleted_list", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var numberingId = EnsureBulletNumbering(document);

            foreach (var item in items)
            {
                var paragraph = new W.Paragraph(
                    new W.ParagraphProperties(
                        new W.NumberingProperties(
                            new W.NumberingLevelReference { Val = 0 },
                            new W.NumberingId { Val = numberingId })),
                    new W.Run(new W.Text(item)));
                body.Append(paragraph);
            }

            document.MainDocumentPart.Document?.Save();
        });
    }

    public void ExcelAddWorksheet(string sessionId, string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Excel);

        ExecuteWriteOperation(session, "excel_add_worksheet", () =>
        {
            using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, true);
            var workbookPart = GetWorkbookPart(spreadsheet);
            var workbook = GetWorkbook(workbookPart);
            var sheets = workbook.Sheets ?? workbook.AppendChild(new Sheets());

            if (GetSheets(workbookPart).Any(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException($"Sheet '{sheetName}' already exists.");
            }

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            worksheetPart.Worksheet.Save();

            var nextSheetId = sheets.Elements<Sheet>().Select(s => s.SheetId?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1;
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = nextSheetId,
                Name = sheetName
            });

            workbook.Save();
        });
    }

    public void PowerPointAddBulletSlide(string sessionId, string title, string bulletLines)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(bulletLines);
        var body = string.Join(Environment.NewLine, bulletLines.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(line => $"- {line}"));
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        ExecuteWriteOperation(session, "powerpoint_add_bullet_slide", () => PowerPointAddSlideCore(session.FilePath, title, body));
    }

    public string BatchExecute(string sessionId, string operationsJson)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(operationsJson);
        _ = GetSession(sessionId);

        var root = JsonNode.Parse(operationsJson) ?? throw new InvalidOperationException("Invalid operations JSON.");
        var operations = root.AsArray();
        var executed = 0;
        var failures = new List<object>();

        foreach (var op in operations)
        {
            if (op is null)
            {
                continue;
            }

            var operation = op["operation"]?.GetValue<string>() ?? string.Empty;
            try
            {
                ExecuteBatchOperation(sessionId, operation, op);
                executed++;
            }
            catch (Exception ex)
            {
                failures.Add(new { operation, error = ex.Message });
            }
        }

        return JsonSerializer.Serialize(new
        {
            total = operations.Count,
            executed,
            failed = failures.Count,
            failures
        });
    }

    public void WordAppendParagraph(string sessionId, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);
        ExecuteWriteOperation(session, "word_append_paragraph", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body;
            if (body is null)
            {
                throw new InvalidOperationException("Word document body is missing.");
            }

            body.AppendChild(new W.Paragraph(new W.Run(new W.Text(text))));
            document.MainDocumentPart!.Document!.Save();
        });
    }

    public void ExcelSetCellValue(string sessionId, string sheetName, string cellReference, string value)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Excel);
        ExecuteWriteOperation(session, "excel_set_cell_value", () =>
        {
            using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, true);
            var workbookPart = GetWorkbookPart(spreadsheet);
            var sheet = GetSheetByName(workbookPart, sheetName);
            var worksheetPart = GetWorksheetPart(workbookPart, sheet);
            var worksheet = GetWorksheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            var rowIndex = ParseRowIndex(cellReference);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
            if (row is null)
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            var cell = row.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));
            if (cell is null)
            {
                cell = new Cell { CellReference = cellReference.ToUpperInvariant() };
                row.Append(cell);
            }

            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(value);

            worksheet.Save();
            GetWorkbook(workbookPart).Save();
        });
    }

    public string ExcelGetCellValue(string sessionId, string sheetName, string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Excel)
        {
            throw new InvalidOperationException("Session document type is not Excel.");
        }

        using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheet = GetSheetByName(workbookPart, sheetName);
        var worksheetPart = GetWorksheetPart(workbookPart, sheet);
        var worksheet = GetWorksheet(worksheetPart);
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));

        return cell?.CellValue?.Text ?? string.Empty;
    }

    public string ExcelGetUsedRange(string sessionId, string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        var session = GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Excel)
        {
            throw new InvalidOperationException("Session document type is not Excel.");
        }

        using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheet = GetSheetByName(workbookPart, sheetName);
        var worksheet = GetWorksheet(GetWorksheetPart(workbookPart, sheet));

        var cells = worksheet.Descendants<Cell>()
            .Where(c => !string.IsNullOrWhiteSpace(c.CellReference?.Value))
            .Select(c => ParseCellReference(c.CellReference!.Value!))
            .ToList();

        if (cells.Count == 0)
        {
            return JsonSerializer.Serialize(new { startCell = string.Empty, endCell = string.Empty, rowCount = 0, columnCount = 0 });
        }

        var minRow = cells.Min(c => c.Row);
        var maxRow = cells.Max(c => c.Row);
        var minCol = cells.Min(c => c.Column);
        var maxCol = cells.Max(c => c.Column);

        return JsonSerializer.Serialize(new
        {
            startCell = BuildCellReference(minCol, minRow),
            endCell = BuildCellReference(maxCol, maxRow),
            rowCount = (int)(maxRow - minRow + 1),
            columnCount = (int)(maxCol - minCol + 1)
        });
    }

    public void ExcelSetRangeValues(string sessionId, string sheetName, string startCell, string valuesJson)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(startCell);
        ArgumentException.ThrowIfNullOrWhiteSpace(valuesJson);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Excel);

        ExecuteWriteOperation(session, "excel_set_range_values", () =>
        {
            var matrixNode = JsonNode.Parse(valuesJson) ?? throw new InvalidOperationException("Invalid valuesJson.");
            var rows = matrixNode.AsArray();
            if (rows.Count == 0)
            {
                throw new InvalidOperationException("valuesJson must contain at least one row.");
            }

            var start = ParseCellReference(startCell);

            using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, true);
            var workbookPart = GetWorkbookPart(spreadsheet);
            var sheet = GetSheetByName(workbookPart, sheetName);
            var worksheetPart = GetWorksheetPart(workbookPart, sheet);
            var worksheet = GetWorksheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            for (var r = 0; r < rows.Count; r++)
            {
                var rowValues = rows[r]?.AsArray() ?? throw new InvalidOperationException("Each row must be an array.");
                var rowIndex = start.Row + (uint)r;
                var row = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex?.Value == rowIndex);
                if (row is null)
                {
                    row = new Row { RowIndex = rowIndex };
                    sheetData.Append(row);
                }

                for (var c = 0; c < rowValues.Count; c++)
                {
                    var value = rowValues[c]?.ToString() ?? string.Empty;
                    var columnIndex = start.Column + c;
                    var cellRef = BuildCellReference(columnIndex, rowIndex);

                    var cell = row.Elements<Cell>().FirstOrDefault(x => string.Equals(x.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                    if (cell is null)
                    {
                        cell = new Cell { CellReference = cellRef };
                        row.Append(cell);
                    }

                    cell.CellFormula = null;
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(value);
                }
            }

            worksheet.Save();
            GetWorkbook(workbookPart).Save();
        });
    }

    public void ExcelSetFormula(string sessionId, string sheetName, string cellReference, string formula)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        ArgumentException.ThrowIfNullOrWhiteSpace(formula);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Excel);

        ExecuteWriteOperation(session, "excel_set_formula", () =>
        {
            using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, true);
            var workbookPart = GetWorkbookPart(spreadsheet);
            var sheet = GetSheetByName(workbookPart, sheetName);
            var worksheetPart = GetWorksheetPart(workbookPart, sheet);
            var worksheet = GetWorksheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            var rowIndex = ParseRowIndex(cellReference);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
            if (row is null)
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            var normalizedRef = cellReference.ToUpperInvariant();
            var cell = row.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, normalizedRef, StringComparison.OrdinalIgnoreCase));
            if (cell is null)
            {
                cell = new Cell { CellReference = normalizedRef };
                row.Append(cell);
            }

            var normalizedFormula = formula.StartsWith('=') ? formula[1..] : formula;
            cell.CellFormula = new CellFormula(normalizedFormula);
            cell.CellValue = null;
            cell.DataType = null;

            worksheet.Save();
            GetWorkbook(workbookPart).Save();
        });
    }

    public string ExcelGetFormula(string sessionId, string sheetName, string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Excel)
        {
            throw new InvalidOperationException("Session document type is not Excel.");
        }

        using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheet = GetSheetByName(workbookPart, sheetName);
        var worksheet = GetWorksheet(GetWorksheetPart(workbookPart, sheet));
        var normalizedRef = cellReference.ToUpperInvariant();
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, normalizedRef, StringComparison.OrdinalIgnoreCase));

        return cell?.CellFormula?.Text ?? string.Empty;
    }

    public void PowerPointAddSlide(string sessionId, string title, string body)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(body);
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);

        ExecuteWriteOperation(session, "powerpoint_add_slide", () => PowerPointAddSlideCore(session.FilePath, title, body));
    }

    public void PowerPointInsertSlideAt(string sessionId, int index, string title, string body)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(body);
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);

        ExecuteWriteOperation(session, "powerpoint_insert_slide_at", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var presentationPart = GetPresentationPart(presentation);
            var slideIdList = GetOrCreateSlideIdList(presentationPart);
            var slides = slideIdList.Elements<SlideId>().ToList();

            if (index < 1 || index > slides.Count + 1)
            {
                throw new InvalidOperationException($"Index {index} is out of range. Valid range is 1..{slides.Count + 1}.");
            }

            var newSlidePart = presentationPart.AddNewPart<SlidePart>();
            newSlidePart.Slide = BuildSimpleSlide(title, body);
            var relationId = presentationPart.GetIdOfPart(newSlidePart);
            var newSlideIdValue = GetNextSlideId(slideIdList);
            var newSlideId = new SlideId { Id = newSlideIdValue, RelationshipId = relationId };

            if (index == slides.Count + 1)
            {
                slideIdList.Append(newSlideId);
            }
            else
            {
                slides[index - 1].InsertBeforeSelf(newSlideId);
            }

            newSlidePart.Slide.Save();
            presentationPart.Presentation?.Save();
        });
    }

    public void PowerPointSetSlideTitle(string sessionId, int slideIndex, string title)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        ExecuteWriteOperation(session, "powerpoint_set_slide_title", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
            SetSlideTextBySlot(slidePart, 0, title);
            GetSlide(slidePart).Save();
        });
    }

    public void PowerPointSetSlideBody(string sessionId, int slideIndex, string body)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(body);
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        ExecuteWriteOperation(session, "powerpoint_set_slide_body", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
            SetSlideTextBySlot(slidePart, 1, body);
            GetSlide(slidePart).Save();
        });
    }

    public void PowerPointReorderSlide(string sessionId, int fromIndex, int toIndex)
    {
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        ExecuteWriteOperation(session, "powerpoint_reorder_slide", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slideIdList = GetOrCreateSlideIdList(GetPresentationPart(presentation));
            var slides = slideIdList.Elements<SlideId>().ToList();

            if (fromIndex < 1 || fromIndex > slides.Count || toIndex < 1 || toIndex > slides.Count)
            {
                throw new InvalidOperationException($"Slide indices must be between 1 and {slides.Count}.");
            }

            if (fromIndex == toIndex)
            {
                return;
            }

            var moved = slides[fromIndex - 1];
            moved.Remove();
            slides.RemoveAt(fromIndex - 1);

            if (toIndex - 1 >= slides.Count)
            {
                slideIdList.Append(moved);
            }
            else
            {
                slides[toIndex - 1].InsertBeforeSelf(moved);
            }

            GetPresentationPart(presentation).Presentation?.Save();
        });
    }

    public void PowerPointDeleteSlide(string sessionId, int slideIndex)
    {
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        ExecuteWriteOperation(session, "powerpoint_delete_slide", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var presentationPart = GetPresentationPart(presentation);
            var slideIdList = GetOrCreateSlideIdList(presentationPart);
            var slides = slideIdList.Elements<SlideId>().ToList();

            if (slideIndex < 1 || slideIndex > slides.Count)
            {
                throw new InvalidOperationException($"Slide index {slideIndex} is out of range. Valid range is 1..{slides.Count}.");
            }

            var target = slides[slideIndex - 1];
            var slidePart = GetSlidePart(presentationPart, target);
            target.Remove();
            presentationPart.DeletePart(slidePart);
            presentationPart.Presentation?.Save();
        });
    }

    private static void PowerPointAddSlideCore(string filePath, string title, string body)
    {
        using var presentation = PresentationDocument.Open(filePath, true);
        var presentationPart = presentation.PresentationPart ?? throw new InvalidOperationException("Presentation part is missing.");
        var slideLayoutPart = EnsurePresentationDefaults(presentationPart);
        var presentationDocument = presentationPart.Presentation ?? throw new InvalidOperationException("Presentation document is missing.");
        var slideIdList = presentationDocument.SlideIdList ?? presentationDocument.AppendChild(new SlideIdList());

        var newSlidePart = presentationPart.AddNewPart<SlidePart>();
        _ = newSlidePart.AddPart(slideLayoutPart);
        newSlidePart.Slide = BuildSimpleSlide(title, body);

        var maxSlideId = slideIdList.Elements<SlideId>().Select(x => x.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max();
        var relationId = presentationPart.GetIdOfPart(newSlidePart);
        slideIdList.Append(new SlideId { Id = maxSlideId + 1, RelationshipId = relationId });

        newSlidePart.Slide.Save();
        presentationDocument.Save();
    }

    private static SlideLayoutPart EnsurePresentationDefaults(PresentationPart presentationPart)
    {
        var presentation = presentationPart.Presentation ??= new Presentation();
        presentation.SlideIdList ??= new SlideIdList();
        presentation.SlideSize ??= new SlideSize { Cx = 10080625, Cy = 7559675 };
        presentation.NotesSize ??= new NotesSize { Cx = 7559675, Cy = 10691813 };

        if (presentationPart.PresentationPropertiesPart is null)
        {
            var presPropsPart = presentationPart.AddNewPart<PresentationPropertiesPart>();
            presPropsPart.PresentationProperties = new PresentationProperties();
            presPropsPart.PresentationProperties.Save();
        }

        if (presentationPart.SlideMasterParts.Any())
        {
            return presentationPart.SlideMasterParts.First().SlideLayoutParts.First();
        }

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        var themePart = slideMasterPart.AddNewPart<ThemePart>();

        slideLayoutPart.SlideLayout = BuildDefaultSlideLayout();

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new SlideLayoutIdList(new SlideLayoutId
            {
                Id = 256U,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            }),
            new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));

        WriteDefaultTheme(themePart);
        _ = presentationPart.AddPart(themePart);

        presentation.SlideMasterIdList ??= new SlideMasterIdList();
        presentation.SlideMasterIdList.Append(new SlideMasterId
        {
            Id = 2147483648U,
            RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
        });

        slideLayoutPart.SlideLayout.Save();
        slideMasterPart.SlideMaster.Save();
        presentation.Save();

        return slideLayoutPart;
    }

    private static SlideLayout BuildDefaultSlideLayout()
    {
        var shapeTree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L },
                new A.ChildOffset { X = 0L, Y = 0L },
                new A.ChildExtents { Cx = 0L, Cy = 0L })));

        shapeTree.Append(CreateLayoutPlaceholderShape(
            id: 2U,
            name: "Title Placeholder",
            placeholderType: PlaceholderValues.Title,
            x: 685800L,
            y: 304800L,
            cx: 9144000L,
            cy: 1143000L));

        shapeTree.Append(CreateLayoutPlaceholderShape(
            id: 3U,
            name: "Content Placeholder",
            placeholderType: PlaceholderValues.Body,
            x: 685800L,
            y: 1828800L,
            cx: 9144000L,
            cy: 4114800L));

        return new SlideLayout(
            new CommonSlideData(shapeTree),
            new ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = SlideLayoutValues.Text,
            Preserve = true
        };
    }

    private static P.Shape CreateLayoutPlaceholderShape(
        uint id,
        string name,
        PlaceholderValues placeholderType,
        long x,
        long y,
        long cx,
        long cy)
    {
        return new P.Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = placeholderType })),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.EndParagraphRunProperties())));
    }

    private static void WriteDefaultTheme(ThemePart themePart)
    {
        const string themeXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office\">"
            + "<a:themeElements>"
            + "<a:clrScheme name=\"Default\">"
            + "<a:dk1><a:srgbClr val=\"000000\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1>"
            + "<a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>"
            + "<a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1><a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2>"
            + "<a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3><a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4>"
            + "<a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6>"
            + "<a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink>"
            + "</a:clrScheme>"
            + "<a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Arial\"/></a:majorFont><a:minorFont><a:latin typeface=\"Arial\"/></a:minorFont></a:fontScheme>"
            + "<a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst>"
            + "<a:lnStyleLst><a:ln w=\"6350\"><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"12700\"><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"19050\"><a:prstDash val=\"solid\"/></a:ln></a:lnStyleLst>"
            + "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>"
            + "<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:bgFillStyleLst>"
            + "</a:fmtScheme></a:themeElements></a:theme>";

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(themeXml));
        themePart.FeedData(stream);
    }

    private static Slide BuildSimpleSlide(string title, string body)
    {
        var shapeTree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        shapeTree.Append(CreateTextShape(2U, "Title", title, 457200L));
        shapeTree.Append(CreateTextShape(3U, "Body", body, 1828800L));

        return new Slide(new CommonSlideData(shapeTree), new ColorMapOverride(new A.MasterColorMapping()));
    }

    private static Shape CreateTextShape(uint id, string name, string text, long yOffset)
    {
        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 457200L, Y = yOffset },
                    new A.Extents { Cx = 8229600L, Cy = 900000L })),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(text)), new A.EndParagraphRunProperties())));
    }

    private static SlideIdList GetOrCreateSlideIdList(PresentationPart presentationPart)
    {
        var presentation = presentationPart.Presentation ?? throw new InvalidOperationException("Presentation document is missing.");
        return presentation.SlideIdList ?? presentation.AppendChild(new SlideIdList());
    }

    private static uint GetNextSlideId(SlideIdList slideIdList)
    {
        return slideIdList.Elements<SlideId>().Select(x => x.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max() + 1;
    }

    private static SlidePart GetSlidePartByIndex(PresentationPart presentationPart, int slideIndex)
    {
        var slideIds = GetSlideIds(presentationPart).ToList();
        if (slideIndex < 1 || slideIndex > slideIds.Count)
        {
            throw new InvalidOperationException($"Slide index {slideIndex} is out of range. Valid range is 1..{slideIds.Count}.");
        }

        return GetSlidePart(presentationPart, slideIds[slideIndex - 1]);
    }

    private static void SetSlideTextBySlot(SlidePart slidePart, int slot, string text)
    {
        var paragraphs = GetSlide(slidePart)
            .Descendants<P.Shape>()
            .Select(shape => shape.TextBody)
            .Where(textBody => textBody is not null)
            .SelectMany(textBody => textBody!.Elements<A.Paragraph>())
            .ToList();

        if (paragraphs.Count == 0)
        {
            throw new InvalidOperationException("Slide does not contain editable text placeholders.");
        }

        var targetIndex = Math.Clamp(slot, 0, paragraphs.Count - 1);
        var paragraph = paragraphs[targetIndex];
        paragraph.RemoveAllChildren<A.Run>();
        paragraph.RemoveAllChildren<A.Break>();
        paragraph.RemoveAllChildren<A.Field>();
        paragraph.Append(new A.Run(new A.Text(text)));
    }

    private static object ListWordStructure(string filePath)
    {
        using var document = WordprocessingDocument.Open(filePath, false);
        var body = document.MainDocumentPart?.Document?.Body;
        var paragraphCount = body?.Descendants<Paragraph>().Count() ?? 0;
        var tableCount = body?.Descendants<W.Table>().Count() ?? 0;
        return new { paragraphCount, tableCount };
    }

    private static object ListExcelStructure(string filePath)
    {
        using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheetNames = GetSheets(workbookPart).Select(s => s.Name?.Value ?? string.Empty).ToArray();
        return new { sheetCount = sheetNames.Length, sheetNames };
    }

    private static object ListPowerPointStructure(string filePath)
    {
        using var presentation = PresentationDocument.Open(filePath, false);
        var slideCount = GetSlideIds(GetPresentationPart(presentation)).Count();
        return new { slideCount };
    }

    private static object FindTextInWord(string filePath, string query)
    {
        using var document = WordprocessingDocument.Open(filePath, false);
        var paragraphs = document.MainDocumentPart?.Document?.Body?.Descendants<Paragraph>() ?? [];
        var matches = paragraphs
            .Select((p, index) => new { index = index + 1, text = p.InnerText })
            .Where(x => x.text.Contains(query, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        return new { matchCount = matches.Length, matches };
    }

    private static object FindTextInExcel(string filePath, string query)
    {
        using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        var matches = new List<object>();

        foreach (var sheet in GetSheets(workbookPart))
        {
            var worksheetPart = GetWorksheetPart(workbookPart, sheet);
            foreach (var cell in GetWorksheet(worksheetPart).Descendants<Cell>())
            {
                var value = ReadCellText(cell, sharedStrings);
                if (!string.IsNullOrEmpty(value) && value.Contains(query, StringComparison.OrdinalIgnoreCase))
                {
                    matches.Add(new { sheet = sheet.Name?.Value ?? string.Empty, cell = cell.CellReference?.Value ?? string.Empty, value });
                }
            }
        }

        return new { matchCount = matches.Count, matches };
    }

    private static object FindTextInPowerPoint(string filePath, string query)
    {
        using var presentation = PresentationDocument.Open(filePath, false);
        var presentationPart = GetPresentationPart(presentation);
        var slideIds = GetSlideIds(presentationPart);
        var matches = new List<object>();

        var index = 0;
        foreach (var slideId in slideIds)
        {
            index++;
            var slidePart = GetSlidePart(presentationPart, slideId);
            var textContent = string.Join(" ", GetSlide(slidePart).Descendants<A.Text>().Select(t => t.Text));
            if (textContent.Contains(query, StringComparison.OrdinalIgnoreCase))
            {
                matches.Add(new { slideIndex = index, text = textContent });
            }
        }

        return new { matchCount = matches.Count, matches };
    }

    private static string ReadCellText(Cell cell, SharedStringTable? sharedStrings)
    {
        var rawValue = cell.CellValue?.Text ?? string.Empty;
        if (cell.DataType?.Value == CellValues.SharedString && int.TryParse(rawValue, out var sharedIndex))
        {
            return sharedStrings?.ElementAtOrDefault(sharedIndex)?.InnerText ?? string.Empty;
        }

        return rawValue;
    }

    private static int CountOccurrences(string input, string value, StringComparison comparison)
    {
        if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(value))
        {
            return 0;
        }

        var count = 0;
        var startIndex = 0;
        while (true)
        {
            var index = input.IndexOf(value, startIndex, comparison);
            if (index < 0)
            {
                break;
            }

            count++;
            startIndex = index + Math.Max(1, value.Length);
        }

        return count;
    }

    private static string ReplaceWithComparison(string input, string oldValue, string newValue, StringComparison comparison)
    {
        if (comparison == StringComparison.Ordinal)
        {
            return input.Replace(oldValue, newValue, StringComparison.Ordinal);
        }

        var startIndex = 0;
        var result = new System.Text.StringBuilder();
        while (true)
        {
            var index = input.IndexOf(oldValue, startIndex, comparison);
            if (index < 0)
            {
                result.Append(input.AsSpan(startIndex));
                break;
            }

            result.Append(input.AsSpan(startIndex, index - startIndex));
            result.Append(newValue);
            startIndex = index + oldValue.Length;
        }

        return result.ToString();
    }

    private static int EnsureBulletNumbering(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
        var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new W.Numbering();

        var numbering = numberingPart.Numbering;
        var existingNumbering = numbering.Elements<W.NumberingInstance>().FirstOrDefault();
        if (existingNumbering?.NumberID?.Value is int existingId)
        {
            return existingId;
        }

        var abstractNumberingId = 1;
        var numberingId = 1;

        numbering.Append(new W.AbstractNum(
                new W.Level(
                    new W.NumberingFormat { Val = W.NumberFormatValues.Bullet },
                    new W.LevelText { Val = "•" },
                    new W.StartNumberingValue { Val = 1 })
                { LevelIndex = 0 })
            { AbstractNumberId = abstractNumberingId });

        numbering.Append(new W.NumberingInstance(new W.AbstractNumId { Val = abstractNumberingId })
        { NumberID = numberingId });

        numbering.Save();
        return numberingId;
    }

    private void ExecuteBatchOperation(string sessionId, string operation, JsonNode payload)
    {
        switch (operation.Trim().ToLowerInvariant())
        {
            case "word_append_paragraph":
                WordAppendParagraph(sessionId, payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'."));
                break;
            case "word_add_table":
                WordAddTable(sessionId, payload["rows"]?.GetValue<int>() ?? 0, payload["columns"]?.GetValue<int>() ?? 0);
                break;
            case "word_insert_paragraph_at":
                WordInsertParagraphAt(
                    sessionId,
                    payload["index"]?.GetValue<int>() ?? 0,
                    payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'."));
                break;
            case "word_replace_text":
                WordReplaceText(
                    sessionId,
                    payload["find"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'find'."),
                    payload["replace"]?.GetValue<string>() ?? string.Empty,
                    payload["matchCase"]?.GetValue<bool>() ?? false);
                break;
            case "word_add_heading":
                WordAddHeading(
                    sessionId,
                    payload["level"]?.GetValue<int>() ?? 1,
                    payload["text"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'text'."));
                break;
            case "word_add_bulleted_list":
                WordAddBulletedList(
                    sessionId,
                    payload["lines"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'lines'."));
                break;
            case "excel_set_cell_value":
                ExcelSetCellValue(
                    sessionId,
                    payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."),
                    payload["cellReference"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'cellReference'."),
                    payload["value"]?.GetValue<string>() ?? string.Empty);
                break;
            case "excel_set_range_values":
                ExcelSetRangeValues(
                    sessionId,
                    payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."),
                    payload["startCell"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'startCell'."),
                    payload["valuesJson"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'valuesJson'."));
                break;
            case "excel_set_formula":
                ExcelSetFormula(
                    sessionId,
                    payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."),
                    payload["cellReference"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'cellReference'."),
                    payload["formula"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'formula'."));
                break;
            case "excel_add_worksheet":
                ExcelAddWorksheet(sessionId, payload["sheetName"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'sheetName'."));
                break;
            case "powerpoint_add_slide":
                PowerPointAddSlide(
                    sessionId,
                    payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."),
                    payload["body"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'body'."));
                break;
            case "powerpoint_insert_slide_at":
            case "power_point_insert_slide_at":
                PowerPointInsertSlideAt(
                    sessionId,
                    payload["index"]?.GetValue<int>() ?? 0,
                    payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."),
                    payload["body"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'body'."));
                break;
            case "powerpoint_set_slide_title":
            case "power_point_set_slide_title":
                PowerPointSetSlideTitle(
                    sessionId,
                    payload["slideIndex"]?.GetValue<int>() ?? 0,
                    payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."));
                break;
            case "powerpoint_set_slide_body":
            case "power_point_set_slide_body":
                PowerPointSetSlideBody(
                    sessionId,
                    payload["slideIndex"]?.GetValue<int>() ?? 0,
                    payload["body"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'body'."));
                break;
            case "powerpoint_reorder_slide":
            case "power_point_reorder_slide":
                PowerPointReorderSlide(
                    sessionId,
                    payload["fromIndex"]?.GetValue<int>() ?? 0,
                    payload["toIndex"]?.GetValue<int>() ?? 0);
                break;
            case "powerpoint_delete_slide":
            case "power_point_delete_slide":
                PowerPointDeleteSlide(
                    sessionId,
                    payload["slideIndex"]?.GetValue<int>() ?? 0);
                break;
            case "powerpoint_add_bullet_slide":
                PowerPointAddBulletSlide(
                    sessionId,
                    payload["title"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'title'."),
                    payload["bulletLines"]?.GetValue<string>() ?? throw new InvalidOperationException("Missing 'bulletLines'."));
                break;
            default:
                throw new InvalidOperationException($"Unsupported batch operation '{operation}'.");
        }
    }

    private OfficeSession GetSession(string sessionId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sessionId);
        if (_sessions.TryGetValue(sessionId, out var session))
        {
            return session;
        }

        throw new InvalidOperationException($"Session '{sessionId}' was not found.");
    }

    private static WorkbookPart GetWorkbookPart(SpreadsheetDocument spreadsheet)
    {
        return spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
    }

    private static IEnumerable<Sheet> GetSheets(WorkbookPart workbookPart)
    {
        var sheets = GetWorkbook(workbookPart).Sheets;
        if (sheets is null)
        {
            throw new InvalidOperationException("Workbook sheets are missing.");
        }

        return sheets.Elements<Sheet>();
    }

    private static Sheet GetSheetByName(WorkbookPart workbookPart, string sheetName)
    {
        return GetSheets(workbookPart).FirstOrDefault(s => s.Name?.Value == sheetName)
            ?? throw new InvalidOperationException($"Sheet '{sheetName}' was not found.");
    }

    private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, Sheet sheet)
    {
        var relationshipId = sheet.Id?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            throw new InvalidOperationException("Sheet relationship id is missing.");
        }

        return (WorksheetPart)workbookPart.GetPartById(relationshipId);
    }

    private static Workbook GetWorkbook(WorkbookPart workbookPart)
    {
        return workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
    }

    private static Worksheet GetWorksheet(WorksheetPart worksheetPart)
    {
        return worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
    }

    private static PresentationPart GetPresentationPart(PresentationDocument presentation)
    {
        return presentation.PresentationPart ?? throw new InvalidOperationException("Presentation part is missing.");
    }

    private static IEnumerable<SlideId> GetSlideIds(PresentationPart presentationPart)
    {
        return presentationPart.Presentation?.SlideIdList?.Elements<SlideId>() ?? [];
    }

    private static SlidePart GetSlidePart(PresentationPart presentationPart, SlideId slideId)
    {
        var relationshipId = slideId.RelationshipId?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            throw new InvalidOperationException("Slide relationship id is missing.");
        }

        return (SlidePart)presentationPart.GetPartById(relationshipId);
    }

    private static Slide GetSlide(SlidePart slidePart)
    {
        return slidePart.Slide ?? throw new InvalidOperationException("Slide content is missing.");
    }

    private OfficeSession GetWritableSession(string sessionId, OfficeDocumentType requiredType)
    {
        var session = GetSession(sessionId);
        if (session.IsReadOnly)
        {
            throw new InvalidOperationException("Session is read-only.");
        }

        if (session.DocumentType != requiredType)
        {
            throw new InvalidOperationException($"Session document type must be '{requiredType}'.");
        }

        return session;
    }

    private static void CreateEmptyDocument(string filePath, OfficeDocumentType documentType)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }

        switch (documentType)
        {
            case OfficeDocumentType.Word:
                using (var word = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                {
                    var mainPart = word.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());
                }
                break;
            case OfficeDocumentType.Excel:
                using (var excel = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = excel.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    sheets.Append(new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1U,
                        Name = "Sheet1"
                    });
                    workbookPart.Workbook.Save();
                }
                break;
            case OfficeDocumentType.PowerPoint:
                InitializePowerPointFromTemplate(filePath);
                break;
            default:
                throw new InvalidOperationException("Unsupported document type.");
        }
    }

    private static void InitializePowerPointFromTemplate(string filePath)
    {
        using (var presentation = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
        {
            var presentationPart = presentation.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            _ = EnsurePresentationDefaults(presentationPart);
            presentationPart.Presentation.Save();
        }
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

    private static OfficeDocumentType GetDocumentTypeFromPath(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension switch
        {
            ".docx" => OfficeDocumentType.Word,
            ".xlsx" => OfficeDocumentType.Excel,
            ".pptx" => OfficeDocumentType.PowerPoint,
            _ => throw new InvalidOperationException($"Unsupported extension '{extension}'.")
        };
    }

    private static string NormalizePath(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        return Path.GetFullPath(filePath);
    }

    private void ExecuteWriteOperation(OfficeSession session, string operationName, Action operation)
    {
        var snapshotPath = CreateSnapshot(session);

        try
        {
            operation();
            AppendOperationLog(session.Id, operationName, "Operation completed.");
        }
        catch
        {
            if (_sessionSnapshots.TryGetValue(session.Id, out var snapshots) && snapshots.Count > 0 && snapshots.Peek() == snapshotPath)
            {
                snapshots.Pop();
            }

            DeleteIfExists(snapshotPath);
            throw;
        }
    }

    private string CreateSnapshot(OfficeSession session)
    {
        var snapshotPath = Path.Combine(Path.GetTempPath(), $"openxmlmcp-snapshot-{session.Id}-{Guid.NewGuid():N}{Path.GetExtension(session.FilePath)}");
        File.Copy(session.FilePath, snapshotPath, overwrite: true);

        var snapshots = _sessionSnapshots.GetOrAdd(session.Id, _ => new Stack<string>());
        snapshots.Push(snapshotPath);
        AppendOperationLog(session.Id, "checkpoint", $"Snapshot created: {snapshotPath}");
        return snapshotPath;
    }

    private void AppendOperationLog(string sessionId, string operationName, string message)
    {
        var history = _sessionOperationLog.GetOrAdd(sessionId, _ => new List<OperationLogEntry>());
        history.Add(new OperationLogEntry(DateTimeOffset.UtcNow, operationName, message));
    }

    private static void DeleteIfExists(string filePath)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }

    private static uint ParseRowIndex(string cellReference)
    {
        var rowText = new string(cellReference.Where(char.IsDigit).ToArray());
        if (uint.TryParse(rowText, out var rowIndex) && rowIndex > 0)
        {
            return rowIndex;
        }

        throw new InvalidOperationException($"Invalid cell reference '{cellReference}'.");
    }

    private static (int Column, uint Row) ParseCellReference(string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);

        var letters = new string(cellReference.Where(char.IsLetter).ToArray()).ToUpperInvariant();
        var numbers = new string(cellReference.Where(char.IsDigit).ToArray());
        if (letters.Length == 0 || !uint.TryParse(numbers, out var row) || row == 0)
        {
            throw new InvalidOperationException($"Invalid cell reference '{cellReference}'.");
        }

        var column = 0;
        foreach (var ch in letters)
        {
            column = (column * 26) + (ch - 'A' + 1);
        }

        return (column, row);
    }

    private static string BuildCellReference(int columnIndex, uint rowIndex)
    {
        if (columnIndex <= 0 || rowIndex == 0)
        {
            throw new InvalidOperationException("Invalid row or column index.");
        }

        var column = new System.Text.StringBuilder();
        var current = columnIndex;
        while (current > 0)
        {
            current--;
            column.Insert(0, (char)('A' + (current % 26)));
            current /= 26;
        }

        return $"{column}{rowIndex}";
    }

    private sealed record OperationLogEntry(DateTimeOffset TimestampUtc, string OperationName, string Message);
}
