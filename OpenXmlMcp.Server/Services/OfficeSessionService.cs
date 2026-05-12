using System.Collections.Concurrent;
using System.Text.Json;
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

    public string OpenDocument(string filePath, bool readOnly = false)
    {
        var normalizedPath = NormalizePath(filePath);

        if (!File.Exists(normalizedPath))
        {
            throw new FileNotFoundException("Office file not found.", normalizedPath);
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
        if (!_sessions.TryRemove(sessionId, out _))
        {
            throw new InvalidOperationException($"Session '{sessionId}' was not found.");
        }
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
        var requiresWrite = normalized is "word_append_paragraph" or "excel_set_cell_value" or "powerpoint_add_slide";
        var expectedType = normalized switch
        {
            "word_append_paragraph" => OfficeDocumentType.Word,
            "excel_set_cell_value" or "excel_get_cell_value" => OfficeDocumentType.Excel,
            "powerpoint_add_slide" => OfficeDocumentType.PowerPoint,
            _ => (OfficeDocumentType?)null
        };

        var warnings = new List<string>();
        var isValid = true;

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

    public void WordAppendParagraph(string sessionId, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Word);

        using var document = WordprocessingDocument.Open(session.FilePath, true);
        var body = document.MainDocumentPart?.Document?.Body;
        if (body is null)
        {
            throw new InvalidOperationException("Word document body is missing.");
        }

        body.AppendChild(new W.Paragraph(new W.Run(new W.Text(text))));
        document.MainDocumentPart!.Document!.Save();
    }

    public void ExcelSetCellValue(string sessionId, string sheetName, string cellReference, string value)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);
        var session = GetWritableSession(sessionId, OfficeDocumentType.Excel);

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

    public void PowerPointAddSlide(string sessionId, string title, string body)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(body);
        var session = GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);

        using var presentation = PresentationDocument.Open(session.FilePath, true);
        var presentationPart = presentation.PresentationPart ?? throw new InvalidOperationException("Presentation part is missing.");
        var presentationDocument = presentationPart.Presentation ?? throw new InvalidOperationException("Presentation document is missing.");
        var slideIdList = presentationDocument.SlideIdList ?? presentationDocument.AppendChild(new SlideIdList());

        var newSlidePart = presentationPart.AddNewPart<SlidePart>();
        newSlidePart.Slide = BuildSimpleSlide(title, body);

        var maxSlideId = slideIdList.Elements<SlideId>().Select(x => x.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max();
        var relationId = presentationPart.GetIdOfPart(newSlidePart);
        slideIdList.Append(new SlideId { Id = maxSlideId + 1, RelationshipId = relationId });

        newSlidePart.Slide.Save();
        presentationDocument.Save();
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
                using (var presentation = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
                {
                    var presentationPart = presentation.AddPresentationPart();
                    presentationPart.Presentation = new Presentation(new SlideIdList());
                    presentationPart.Presentation.Save();
                }
                break;
            default:
                throw new InvalidOperationException("Unsupported document type.");
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

    private static uint ParseRowIndex(string cellReference)
    {
        var rowText = new string(cellReference.Where(char.IsDigit).ToArray());
        if (uint.TryParse(rowText, out var rowIndex) && rowIndex > 0)
        {
            return rowIndex;
        }

        throw new InvalidOperationException($"Invalid cell reference '{cellReference}'.");
    }
}
