using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlMcp.Server.Models;
using DocumentFormat.OpenXml;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;

namespace OpenXmlMcp.Server.Services;

/// <summary>
/// Handles all Excel (XLSX) document operations. Depends on <see cref="SessionManager"/>
/// for session access and write orchestration.
/// </summary>
public class ExcelDocumentService(SessionManager sessionManager)
{
    private const string GeneratedByApplication = "OpenXmlMcp";
    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    public string SetCellValue(string sessionId, string sheetName, string cellReference, string value)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Excel);
        sessionManager.ExecuteWriteOperation(session, "excel_set_cell_value", () =>
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
            MarkWorkbookForRecalculation(workbookPart);
            GetWorkbook(workbookPart).Save();
        });

        return SessionManager.BuildMutationResult("excel_set_cell_value", new { sheetName, cellReference = cellReference.ToUpperInvariant() });
    }

    public string SetCellStyle(string sessionId, string sheetName, string cellReference, string fontName, int fontSize, bool bold, bool italic, string colorHex)
    {
        SetCellStyleInternal(sessionId, sheetName, cellReference, new TextStyle(fontName, fontSize, bold, italic, colorHex));
        return SessionManager.BuildMutationResult("excel_set_cell_style", new { sheetName, cellReference = cellReference.ToUpperInvariant() });
    }

    public string GetCellValue(string sessionId, string sheetName, string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = sessionManager.GetSession(sessionId);
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

    public string GetCellInfo(string sessionId, string sheetName, string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = sessionManager.GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Excel)
        {
            throw new InvalidOperationException("Session document type is not Excel.");
        }

        using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheet = GetSheetByName(workbookPart, sheetName);
        var worksheetPart = GetWorksheetPart(workbookPart, sheet);
        var worksheet = GetWorksheet(worksheetPart);
        var normalizedRef = cellReference.ToUpperInvariant();
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, normalizedRef, StringComparison.OrdinalIgnoreCase));

        if (cell is null)
        {
            return JsonSerializer.Serialize(new
            {
                sheetName,
                cellReference = normalizedRef,
                exists = false,
                value = string.Empty,
                formula = string.Empty,
                cachedValue = string.Empty,
                dataType = string.Empty,
                styleIndex = -1L
            });
        }

        var styleDetails = ResolveExcelCellStyle(workbookPart, cell);

        return JsonSerializer.Serialize(new
        {
            sheetName,
            cellReference = normalizedRef,
            exists = true,
            value = cell.CellValue?.Text ?? string.Empty,
            formula = cell.CellFormula?.Text ?? string.Empty,
            cachedValue = cell.CellValue?.Text ?? string.Empty,
            dataType = cell.DataType?.Value.ToString() ?? string.Empty,
            styleIndex = cell.StyleIndex?.Value ?? 0UL,
            style = styleDetails
        });
    }

    public string GetCellStyle(string sessionId, string sheetName, string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = sessionManager.GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Excel)
        {
            throw new InvalidOperationException("Session document type is not Excel.");
        }

        using var spreadsheet = SpreadsheetDocument.Open(session.FilePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheet = GetSheetByName(workbookPart, sheetName);
        var worksheetPart = GetWorksheetPart(workbookPart, sheet);
        var worksheet = GetWorksheet(worksheetPart);
        var normalizedRef = cellReference.ToUpperInvariant();
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, normalizedRef, StringComparison.OrdinalIgnoreCase));

        if (cell is null)
        {
            return JsonSerializer.Serialize(new { sheetName, cellReference = normalizedRef, exists = false, styleIndex = -1, style = new { } });
        }

        return JsonSerializer.Serialize(new
        {
            sheetName,
            cellReference = normalizedRef,
            exists = true,
            styleIndex = (long)(cell.StyleIndex?.Value ?? 0UL),
            style = ResolveExcelCellStyle(workbookPart, cell)
        });
    }

    public string GetUsedRange(string sessionId, string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        var session = sessionManager.GetSession(sessionId);
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

    public string SetRangeValues(string sessionId, string sheetName, string startCell, string valuesJson)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(startCell);
        ArgumentException.ThrowIfNullOrWhiteSpace(valuesJson);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Excel);

        sessionManager.ExecuteWriteOperation(session, "excel_set_range_values", () =>
        {
            JsonNode matrixNode;
            try
            {
                matrixNode = JsonNode.Parse(valuesJson) ?? throw new InvalidOperationException("Invalid valuesJson.");
            }
            catch (System.Text.Json.JsonException ex)
            {
                throw new InvalidOperationException("Invalid valuesJson. Provide strict JSON with double-quoted strings, e.g. [[\"A\",\"B\"],[\"C\",\"D\"]].", ex);
            }

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
                    var node = rowValues[c];
                    var columnIndex = start.Column + c;
                    var cellRef = BuildCellReference(columnIndex, rowIndex);

                    var cell = row.Elements<Cell>().FirstOrDefault(x => string.Equals(x.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                    if (cell is null)
                    {
                        cell = new Cell { CellReference = cellRef };
                        row.Append(cell);
                    }

                    ApplyExcelNodeValue(cell, node);
                }
            }

            worksheet.Save();
            GetWorkbook(workbookPart).Save();
        });

        return SessionManager.BuildMutationResult("excel_set_range_values", new { sheetName, startCell = startCell.ToUpperInvariant() });
    }

    public string SetFormula(string sessionId, string sheetName, string cellReference, string formula)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        ArgumentException.ThrowIfNullOrWhiteSpace(formula);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Excel);

        sessionManager.ExecuteWriteOperation(session, "excel_set_formula", () =>
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

        return SessionManager.BuildMutationResult("excel_set_formula", new { sheetName, cellReference = cellReference.ToUpperInvariant() });
    }

    public string GetFormula(string sessionId, string sheetName, string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = sessionManager.GetSession(sessionId);
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

    public string AddWorksheet(string sessionId, string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Excel);

        sessionManager.ExecuteWriteOperation(session, "excel_add_worksheet", () =>
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

        return SessionManager.BuildMutationResult("excel_add_worksheet", new { sheetName });
    }

    // -------------------------------------------------------------------------
    // Internal helpers used by OfficeSessionService for cross-cutting features
    // -------------------------------------------------------------------------

    public void SetCellStyleInternal(string sessionId, string sheetName, string cellReference, TextStyle style)
    {
        WordDocumentService.ValidateTextStyle(style);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Excel);
        sessionManager.ExecuteWriteOperation(session, "excel_set_cell_style", () =>
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
                cell = new Cell { CellReference = normalizedRef, DataType = CellValues.String, CellValue = new CellValue(string.Empty) };
                row.Append(cell);
            }

            var styleIndex = EnsureExcelCellFormat(workbookPart, style);
            cell.StyleIndex = styleIndex;

            worksheet.Save();
            MarkWorkbookForRecalculation(workbookPart);
            GetWorkbook(workbookPart).Save();
        });
    }

    public void ApplyStylePreset(string filePath)
    {
        using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet ??= new Stylesheet(
            new DocumentFormat.OpenXml.Spreadsheet.Fonts(
                new DocumentFormat.OpenXml.Spreadsheet.Font(
                    new FontName { Val = "Calibri" },
                    new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 11D }))
            { Count = 1U },
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
            { Count = 2U },
            new Borders(new DocumentFormat.OpenXml.Spreadsheet.Border()) { Count = 1U },
            new CellStyleFormats(new CellFormat()) { Count = 1U },
            new CellFormats(new CellFormat()) { Count = 1U });

        stylesPart.Stylesheet.Save();
        GetWorkbook(workbookPart).Save();
    }

    public void InitializeEmptyDocument(string filePath, string? creator = null)
    {
        using var excel = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        if (!string.IsNullOrWhiteSpace(creator))
        {
            excel.PackageProperties.Creator = creator;
            excel.PackageProperties.LastModifiedBy = creator;
        }
        var appPart = excel.ExtendedFilePropertiesPart ?? excel.AddNewPart<ExtendedFilePropertiesPart>();
        appPart.Properties ??= new Ap.Properties();
        appPart.Properties.Application = new Ap.Application(GeneratedByApplication);
        appPart.Properties.Save();
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

    public object ListStructure(string filePath)
    {
        using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = GetWorkbookPart(spreadsheet);
        var sheetNames = GetSheets(workbookPart).Select(s => s.Name?.Value ?? string.Empty).ToArray();
        return new { sheetCount = sheetNames.Length, sheetNames };
    }

    public object FindText(string filePath, string query)
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

    // -------------------------------------------------------------------------
    // Private implementation details
    // -------------------------------------------------------------------------

    private static uint EnsureExcelCellFormat(WorkbookPart workbookPart, TextStyle style)
    {
        var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet ??= new Stylesheet(
            new DocumentFormat.OpenXml.Spreadsheet.Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()) { Count = 1U },
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2U },
            new Borders(new DocumentFormat.OpenXml.Spreadsheet.Border()) { Count = 1U },
            new CellStyleFormats(new CellFormat()) { Count = 1U },
            new CellFormats(new CellFormat()) { Count = 1U });

        var stylesheet = stylesPart.Stylesheet;
        stylesheet.Fonts ??= new DocumentFormat.OpenXml.Spreadsheet.Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()) { Count = 1U };
        stylesheet.CellFormats ??= new CellFormats(new CellFormat()) { Count = 1U };

        var font = new DocumentFormat.OpenXml.Spreadsheet.Font(
            new FontName { Val = style.FontName },
            new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = style.FontSize },
            new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = style.ColorHex });
        if (style.Bold)
        {
            font.Append(new DocumentFormat.OpenXml.Spreadsheet.Bold());
        }

        if (style.Italic)
        {
            font.Append(new DocumentFormat.OpenXml.Spreadsheet.Italic());
        }

        stylesheet.Fonts.Append(font);
        stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

        var fontId = stylesheet.Fonts.Count.Value - 1;
        var format = new CellFormat { FontId = fontId, ApplyFont = true };
        stylesheet.CellFormats.Append(format);
        stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

        stylesheet.Save();
        return stylesheet.CellFormats.Count.Value - 1;
    }

    private static object ResolveExcelCellStyle(WorkbookPart workbookPart, Cell cell)
    {
        var styleIndex = cell.StyleIndex?.Value ?? 0UL;
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        var cellFormat = stylesheet?.CellFormats?.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
        var fontId = (int)(cellFormat?.FontId?.Value ?? 0U);
        var font = stylesheet?.Fonts?.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAtOrDefault(fontId);

        return new
        {
            fontName = font?.FontName?.Val?.Value ?? string.Empty,
            fontSize = font?.FontSize?.Val?.Value ?? 0,
            bold = font?.Bold is not null,
            italic = font?.Italic is not null,
            colorHex = font?.Color?.Rgb?.Value?.Replace("FF", string.Empty, StringComparison.OrdinalIgnoreCase) ?? string.Empty
        };
    }

    private static void MarkWorkbookForRecalculation(WorkbookPart workbookPart)
    {
        var workbook = GetWorkbook(workbookPart);
        workbook.CalculationProperties ??= new CalculationProperties();
        workbook.CalculationProperties.FullCalculationOnLoad = true;
        workbook.CalculationProperties.ForceFullCalculation = true;
    }

    private static void ApplyExcelNodeValue(Cell cell, JsonNode? node)
    {
        if (node is null)
        {
            cell.CellFormula = null;
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(string.Empty);
            return;
        }

        if (node is JsonValue valueNode)
        {
            if (valueNode.TryGetValue<string>(out var stringValue))
            {
                if (!string.IsNullOrEmpty(stringValue) && stringValue.StartsWith("=", StringComparison.Ordinal))
                {
                    cell.CellFormula = new CellFormula(stringValue[1..]);
                    cell.CellValue = null;
                    cell.DataType = null;
                    return;
                }

                cell.CellFormula = null;
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(stringValue);
                return;
            }

            if (valueNode.TryGetValue<double>(out var numberValue))
            {
                cell.CellFormula = null;
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(numberValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
                return;
            }

            if (valueNode.TryGetValue<bool>(out var boolValue))
            {
                cell.CellFormula = null;
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(boolValue ? "1" : "0");
                return;
            }
        }

        cell.CellFormula = null;
        cell.DataType = CellValues.String;
        cell.CellValue = new CellValue(node.ToJsonString());
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

    private static WorkbookPart GetWorkbookPart(SpreadsheetDocument spreadsheet)
        => spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");

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
        => GetSheets(workbookPart).FirstOrDefault(s => s.Name?.Value == sheetName)
            ?? throw new InvalidOperationException($"Sheet '{sheetName}' was not found.");

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
        => workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");

    private static Worksheet GetWorksheet(WorksheetPart worksheetPart)
        => worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");

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
}
