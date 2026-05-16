using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlMcp.Server.Models;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMcp.Server.Services;

/// <summary>
/// Handles all Word (DOCX) document operations. Depends on <see cref="SessionManager"/>
/// for session access and write orchestration.
/// </summary>
public class WordDocumentService(SessionManager sessionManager)
{
    private const int DefaultParagraphBeforePt = 0;
    private const int DefaultParagraphAfterPt = 8;
    private const double DefaultParagraphLineSpacing = 1.15;
    private const int DefaultHeadingBeforePt = 12;
    private const int DefaultHeadingAfterPt = 6;
    private const double DefaultHeadingLineSpacing = 1.15;
    private const int DefaultListBeforePt = 0;
    private const int DefaultListAfterPt = 2;
    private const double DefaultListLineSpacing = 1.15;

    private static readonly IReadOnlyList<WordBuiltInStyleDefinition> WordBuiltInStyles = BuildWordBuiltInStyles();

    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    public string AppendParagraph(string sessionId, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var paragraphIndex = -1;
        sessionManager.ExecuteWriteOperation(session, "word_append_paragraph", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body
                ?? throw new InvalidOperationException("Word document body is missing.");
            body.AppendChild(new W.Paragraph(new W.Run(new W.Text(text))));
            var paragraph = body.Elements<W.Paragraph>().Last();
            ApplyWordParagraphSpacing(paragraph, DefaultParagraphBeforePt, DefaultParagraphAfterPt, DefaultParagraphLineSpacing);
            paragraphIndex = body.Elements<W.Paragraph>().Count();
            document.MainDocumentPart!.Document!.Save();
        });

        return SessionManager.BuildMutationResult("word_append_paragraph", new { textPreview = TrimPreview(text), paragraphIndex });
    }

    public string SetParagraphStyle(string sessionId, int paragraphIndex, string fontName, int fontSize, bool bold, bool italic, string colorHex)
    {
        SetParagraphStyle(sessionId, paragraphIndex, new TextStyle(fontName, fontSize, bold, italic, colorHex));
        return SessionManager.BuildMutationResult("word_set_paragraph_style", new { paragraphIndex });
    }

    public string SetParagraphSpacing(string sessionId, int paragraphIndex, int beforePt, int afterPt, double lineSpacing)
    {
        if (beforePt < 0 || afterPt < 0)
        {
            throw new InvalidOperationException("Paragraph spacing before/after must be >= 0.");
        }

        if (lineSpacing <= 0)
        {
            throw new InvalidOperationException("Line spacing must be greater than 0.");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_set_paragraph_spacing", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();
            if (paragraphIndex < 1 || paragraphIndex > paragraphs.Count)
            {
                throw new InvalidOperationException(BuildWordBodyParagraphRangeMessage(body, paragraphIndex, paragraphs.Count));
            }

            ApplyWordParagraphSpacing(paragraphs[paragraphIndex - 1], beforePt, afterPt, lineSpacing);
            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_set_paragraph_spacing", new { paragraphIndex, beforePt, afterPt, lineSpacing });
    }

    public string SetDocumentSpacingPreset(string sessionId, string preset = "normal")
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(preset);
        var normalized = preset.Trim().ToLowerInvariant();
        var settings = normalized switch
        {
            "compact" => (beforePt: 0, afterPt: 4, lineSpacing: 1.0),
            "comfortable" => (beforePt: 0, afterPt: 10, lineSpacing: 1.3),
            _ => (beforePt: DefaultParagraphBeforePt, afterPt: DefaultParagraphAfterPt, lineSpacing: DefaultParagraphLineSpacing)
        };

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_set_document_spacing_preset", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            foreach (var paragraph in body.Elements<W.Paragraph>())
            {
                ApplyWordParagraphSpacing(paragraph, settings.beforePt, settings.afterPt, settings.lineSpacing);
            }

            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_set_document_spacing_preset", new { preset = normalized });
    }

    public string AddTable(string sessionId, int rows, int columns)
    {
        if (rows <= 0 || columns <= 0)
        {
            throw new InvalidOperationException("Rows and columns must be greater than zero.");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var tableIndex = -1;
        sessionManager.ExecuteWriteOperation(session, "word_add_table", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            tableIndex = body.Elements<W.Table>().Count() + 1;

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

        return SessionManager.BuildMutationResult("word_add_table", new { rows, columns, tableIndex });
    }

    public string SetTableCell(string sessionId, int tableIndex, int rowIndex, int columnIndex, string text)
    {
        if (tableIndex < 1 || rowIndex < 1 || columnIndex < 1)
        {
            throw new InvalidOperationException("tableIndex, rowIndex, and columnIndex must be >= 1.");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_set_table_cell", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var table = GetWordTableByIndex(body, tableIndex);
            var row = table.Elements<W.TableRow>().ElementAtOrDefault(rowIndex - 1)
                ?? throw new InvalidOperationException($"Row index {rowIndex} is out of range for table {tableIndex}.");
            var cell = row.Elements<W.TableCell>().ElementAtOrDefault(columnIndex - 1)
                ?? throw new InvalidOperationException($"Column index {columnIndex} is out of range for table {tableIndex}, row {rowIndex}.");

            cell.RemoveAllChildren<W.Paragraph>();
            cell.Append(new W.Paragraph(new W.Run(new W.Text(text ?? string.Empty))));
            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_set_table_cell", new { tableIndex, rowIndex, columnIndex });
    }

    public string GetTableCell(string sessionId, int tableIndex, int rowIndex, int columnIndex)
    {
        if (tableIndex < 1 || rowIndex < 1 || columnIndex < 1)
        {
            throw new InvalidOperationException("tableIndex, rowIndex, and columnIndex must be >= 1.");
        }

        var session = sessionManager.GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Word)
        {
            throw new InvalidOperationException("Session document type is not Word.");
        }

        using var document = WordprocessingDocument.Open(session.FilePath, false);
        var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
        var table = GetWordTableByIndex(body, tableIndex);
        var row = table.Elements<W.TableRow>().ElementAtOrDefault(rowIndex - 1)
            ?? throw new InvalidOperationException($"Row index {rowIndex} is out of range for table {tableIndex}.");
        var cell = row.Elements<W.TableCell>().ElementAtOrDefault(columnIndex - 1)
            ?? throw new InvalidOperationException($"Column index {columnIndex} is out of range for table {tableIndex}, row {rowIndex}.");

        return string.Join("\n", cell.Elements<W.Paragraph>().Select(p => p.InnerText));
    }

    public string GetParagraphInfo(string sessionId, int paragraphIndex)
    {
        var session = sessionManager.GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Word)
        {
            throw new InvalidOperationException("Session document type is not Word.");
        }

        using var document = WordprocessingDocument.Open(session.FilePath, false);
        var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
        var paragraphs = body.Elements<W.Paragraph>().ToList();
        if (paragraphIndex < 1 || paragraphIndex > paragraphs.Count)
        {
            throw new InvalidOperationException(BuildWordBodyParagraphRangeMessage(body, paragraphIndex, paragraphs.Count));
        }

        var paragraph = paragraphs[paragraphIndex - 1];
        var pPr = paragraph.ParagraphProperties;
        var spacing = pPr?.SpacingBetweenLines;
        var styleId = pPr?.ParagraphStyleId?.Val?.Value ?? string.Empty;
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<W.Style>() ?? [];
        var styleName = styles.FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase))?.StyleName?.Val?.Value ?? string.Empty;

        return JsonSerializer.Serialize(new
        {
            paragraphIndex,
            text = paragraph.InnerText,
            styleId,
            styleName,
            spacingBeforeTwips = spacing?.Before?.Value ?? string.Empty,
            spacingAfterTwips = spacing?.After?.Value ?? string.Empty,
            lineTwips = spacing?.Line?.Value ?? string.Empty,
            lineRule = spacing?.LineRule?.Value.ToString() ?? string.Empty
        });
    }

    public string InsertParagraphAt(string sessionId, int index, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);

        sessionManager.ExecuteWriteOperation(session, "word_insert_paragraph_at", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();

            if (index < 1 || index > paragraphs.Count + 1)
            {
                throw new InvalidOperationException($"Index {index} is out of range. Valid range is 1..{paragraphs.Count + 1}.");
            }

            var newParagraph = new W.Paragraph(new W.Run(new W.Text(text)));
            ApplyWordParagraphSpacing(newParagraph, DefaultParagraphBeforePt, DefaultParagraphAfterPt, DefaultParagraphLineSpacing);
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

        return SessionManager.BuildMutationResult("word_insert_paragraph_at", new { index });
    }

    public string ReplaceText(string sessionId, string find, string replace, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(find);
        replace ??= string.Empty;
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var replacements = 0;

        sessionManager.ExecuteWriteOperation(session, "word_replace_text", () =>
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

        return SessionManager.BuildMutationResult("word_replace_text", new { replacementCount = replacements }, replacements > 0);
    }

    public string AddHeading(string sessionId, int level, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        if (level < 1 || level > 9)
        {
            throw new InvalidOperationException("Heading level must be between 1 and 9.");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var paragraphIndex = -1;
        sessionManager.ExecuteWriteOperation(session, "word_add_heading", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");

            var heading = new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = $"Heading{level}" }),
                new W.Run(new W.Text(text)));
            ApplyWordParagraphSpacing(heading, DefaultHeadingBeforePt, DefaultHeadingAfterPt, DefaultHeadingLineSpacing);

            body.Append(heading);
            paragraphIndex = body.Elements<W.Paragraph>().Count();
            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_add_heading", new { level, textPreview = TrimPreview(text), paragraphIndex });
    }

    public string ApplyTableStyle(string sessionId, int tableIndex, string styleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_apply_table_style", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var table = GetWordTableByIndex(body, tableIndex);

            var tableProperties = table.Elements<W.TableProperties>().FirstOrDefault();
            if (tableProperties is null)
            {
                tableProperties = new W.TableProperties();
                table.InsertAt(tableProperties, 0);
            }

            tableProperties.TableStyle = new W.TableStyle { Val = styleName };
            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_apply_table_style", new { tableIndex, styleName });
    }

    public string FormatTableHeaderRow(string sessionId, int tableIndex, bool bold = true, string? shadingFill = null, string? fontName = null, string? colorHex = null)
    {
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_format_table_header_row", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var table = GetWordTableByIndex(body, tableIndex);

            var headerRow = table.Elements<W.TableRow>().FirstOrDefault()
                ?? throw new InvalidOperationException($"Table {tableIndex} has no rows.");

            // Mark the row as a header row
            var rowProps = headerRow.Elements<W.TableRowProperties>().FirstOrDefault();
            if (rowProps is null)
            {
                rowProps = new W.TableRowProperties();
                headerRow.InsertAt(rowProps, 0);
            }

            if (!rowProps.Elements<W.TableHeader>().Any())
            {
                rowProps.Append(new W.TableHeader());
            }

            // Apply shading and font to each cell
            foreach (var cell in headerRow.Elements<W.TableCell>())
            {
                if (!string.IsNullOrWhiteSpace(shadingFill))
                {
                    var cellProps = cell.Elements<W.TableCellProperties>().FirstOrDefault();
                    if (cellProps is null)
                    {
                        cellProps = new W.TableCellProperties();
                        cell.InsertAt(cellProps, 0);
                    }

                    cellProps.Shading = new W.Shading
                    {
                        Val = W.ShadingPatternValues.Clear,
                        Fill = shadingFill.TrimStart('#'),
                        Color = "auto"
                    };
                }

                foreach (var para in cell.Elements<W.Paragraph>())
                {
                    foreach (var run in para.Elements<W.Run>())
                    {
                        run.RunProperties ??= new W.RunProperties();
                        if (bold)
                        {
                            run.RunProperties.Bold = new W.Bold();
                        }

                        if (!string.IsNullOrWhiteSpace(fontName))
                        {
                            run.RunProperties.RunFonts = new W.RunFonts { Ascii = fontName, HighAnsi = fontName, ComplexScript = fontName };
                        }

                        if (!string.IsNullOrWhiteSpace(colorHex))
                        {
                            run.RunProperties.Color = new W.Color { Val = colorHex.TrimStart('#') };
                        }
                    }

                    // Also apply to paragraph mark run properties
                    var pPr = para.ParagraphProperties ??= new W.ParagraphProperties();
                    pPr.ParagraphMarkRunProperties ??= new W.ParagraphMarkRunProperties();
                    if (bold && !pPr.ParagraphMarkRunProperties.Elements<W.Bold>().Any())
                    {
                        pPr.ParagraphMarkRunProperties.Append(new W.Bold());
                    }
                }
            }

            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_format_table_header_row", new { tableIndex, bold, shadingFill, fontName, colorHex });
    }

    public string SetTableValues(string sessionId, int tableIndex, string valuesJson, int startRow = 1, int startColumn = 1)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(valuesJson);
        if (tableIndex < 1) throw new InvalidOperationException("tableIndex must be >= 1.");
        if (startRow < 1) throw new InvalidOperationException("startRow must be >= 1.");
        if (startColumn < 1) throw new InvalidOperationException("startColumn must be >= 1.");

        var rows = JsonSerializer.Deserialize<List<List<JsonElement>>>(valuesJson)
            ?? throw new InvalidOperationException("valuesJson must be a JSON array of arrays.");
        if (rows.Count == 0)
        {
            throw new InvalidOperationException("valuesJson must contain at least one row.");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var updatedCellCount = 0;
        sessionManager.ExecuteWriteOperation(session, "word_set_table_values", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var table = GetWordTableByIndex(body, tableIndex);
            var tableRows = table.Elements<W.TableRow>().ToList();

            for (var r = 0; r < rows.Count; r++)
            {
                var rowData = rows[r];
                var rowIdx = startRow - 1 + r;
                if (rowIdx >= tableRows.Count)
                {
                    throw new InvalidOperationException($"Row {rowIdx + 1} does not exist in table {tableIndex} (table has {tableRows.Count} rows).");
                }

                var tableRow = tableRows[rowIdx];
                var tableCells = tableRow.Elements<W.TableCell>().ToList();

                for (var c = 0; c < rowData.Count; c++)
                {
                    var colIdx = startColumn - 1 + c;
                    if (colIdx >= tableCells.Count)
                    {
                        throw new InvalidOperationException($"Column {colIdx + 1} does not exist in table {tableIndex}, row {rowIdx + 1} (row has {tableCells.Count} columns).");
                    }

                    var cell = tableCells[colIdx];
                    var cellText = rowData[c].ValueKind == JsonValueKind.String
                        ? rowData[c].GetString() ?? string.Empty
                        : rowData[c].ToString();

                    cell.RemoveAllChildren<W.Paragraph>();
                    cell.Append(new W.Paragraph(new W.Run(new W.Text(cellText))));
                    updatedCellCount++;
                }
            }

            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_set_table_values", new { tableIndex, startRow, startColumn, updatedCellCount });
    }

    public string AddTableRow(string sessionId, int tableIndex, int? rowIndex = null)
    {
        if (tableIndex < 1) throw new InvalidOperationException("tableIndex must be >= 1.");

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var insertedRowIndex = -1;
        var rowCount = -1;
        sessionManager.ExecuteWriteOperation(session, "word_add_table_row", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var table = GetWordTableByIndex(body, tableIndex);
            var rows = table.Elements<W.TableRow>().ToList();

            // Derive column count from first row (or 1 if table is empty)
            var columnCount = rows.Count > 0 ? rows[0].Elements<W.TableCell>().Count() : 1;

            var newRow = new W.TableRow();
            for (var c = 0; c < columnCount; c++)
            {
                newRow.Append(new W.TableCell(new W.Paragraph(new W.Run(new W.Text(string.Empty)))));
            }

            if (rowIndex.HasValue)
            {
                if (rowIndex.Value < 1 || rowIndex.Value > rows.Count + 1)
                {
                    throw new InvalidOperationException($"rowIndex {rowIndex.Value} is out of range. Valid range is 1..{rows.Count + 1}.");
                }

                if (rowIndex.Value <= rows.Count)
                {
                    rows[rowIndex.Value - 1].InsertBeforeSelf(newRow);
                    insertedRowIndex = rowIndex.Value;
                }
                else
                {
                    table.Append(newRow);
                    insertedRowIndex = rows.Count + 1;
                }
            }
            else
            {
                table.Append(newRow);
                insertedRowIndex = rows.Count + 1;
            }

            rowCount = table.Elements<W.TableRow>().Count();
            document.MainDocumentPart!.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_add_table_row", new { tableIndex, rowIndex = insertedRowIndex, rowCount });
    }

    public string DeleteTableRow(string sessionId, int tableIndex, int rowIndex)
    {
        if (tableIndex < 1) throw new InvalidOperationException("tableIndex must be >= 1.");
        if (rowIndex < 1) throw new InvalidOperationException("rowIndex must be >= 1.");

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var rowCount = -1;
        sessionManager.ExecuteWriteOperation(session, "word_delete_table_row", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var table = GetWordTableByIndex(body, tableIndex);
            var rows = table.Elements<W.TableRow>().ToList();

            if (rows.Count <= 1)
            {
                throw new InvalidOperationException($"Cannot delete the only row in table {tableIndex}.");
            }

            if (rowIndex > rows.Count)
            {
                throw new InvalidOperationException($"rowIndex {rowIndex} is out of range. Table {tableIndex} has {rows.Count} rows.");
            }

            rows[rowIndex - 1].Remove();
            rowCount = table.Elements<W.TableRow>().Count();
            document.MainDocumentPart!.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_delete_table_row", new { tableIndex, removedIndex = rowIndex, rowCount });
    }

    public string AddTableColumn(string sessionId, int tableIndex, int? columnIndex = null)
    {
        if (tableIndex < 1) throw new InvalidOperationException("tableIndex must be >= 1.");

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var insertedColumnIndex = -1;
        var columnCount = -1;
        sessionManager.ExecuteWriteOperation(session, "word_add_table_column", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var table = GetWordTableByIndex(body, tableIndex);
            var rows = table.Elements<W.TableRow>().ToList();

            if (rows.Count == 0)
            {
                throw new InvalidOperationException($"Table {tableIndex} has no rows.");
            }

            var firstRowCells = rows[0].Elements<W.TableCell>().ToList();
            var currentColumnCount = firstRowCells.Count;

            // Validate / normalise columnIndex
            int insertAt;
            if (columnIndex.HasValue)
            {
                if (columnIndex.Value < 1 || columnIndex.Value > currentColumnCount + 1)
                {
                    throw new InvalidOperationException($"columnIndex {columnIndex.Value} is out of range. Valid range is 1..{currentColumnCount + 1}.");
                }

                insertAt = columnIndex.Value;
            }
            else
            {
                insertAt = currentColumnCount + 1;
            }

            foreach (var row in rows)
            {
                var cells = row.Elements<W.TableCell>().ToList();
                var newCell = new W.TableCell(new W.Paragraph(new W.Run(new W.Text(string.Empty))));

                if (insertAt <= cells.Count)
                {
                    cells[insertAt - 1].InsertBeforeSelf(newCell);
                }
                else
                {
                    row.Append(newCell);
                }
            }

            insertedColumnIndex = insertAt;
            columnCount = rows[0].Elements<W.TableCell>().Count();
            document.MainDocumentPart!.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_add_table_column", new { tableIndex, columnIndex = insertedColumnIndex, columnCount });
    }

    public string DeleteTableColumn(string sessionId, int tableIndex, int columnIndex)
    {
        if (tableIndex < 1) throw new InvalidOperationException("tableIndex must be >= 1.");
        if (columnIndex < 1) throw new InvalidOperationException("columnIndex must be >= 1.");

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var columnCount = -1;
        sessionManager.ExecuteWriteOperation(session, "word_delete_table_column", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var table = GetWordTableByIndex(body, tableIndex);
            var rows = table.Elements<W.TableRow>().ToList();

            if (rows.Count == 0)
            {
                throw new InvalidOperationException($"Table {tableIndex} has no rows.");
            }

            var firstRowCells = rows[0].Elements<W.TableCell>().ToList();
            if (firstRowCells.Count <= 1)
            {
                throw new InvalidOperationException($"Cannot delete the only column in table {tableIndex}.");
            }

            if (columnIndex > firstRowCells.Count)
            {
                throw new InvalidOperationException($"columnIndex {columnIndex} is out of range. Table {tableIndex} has {firstRowCells.Count} columns.");
            }

            foreach (var row in rows)
            {
                var cells = row.Elements<W.TableCell>().ToList();
                if (columnIndex <= cells.Count)
                {
                    cells[columnIndex - 1].Remove();
                }
            }

            columnCount = rows[0].Elements<W.TableCell>().Count();
            document.MainDocumentPart!.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_delete_table_column", new { tableIndex, removedIndex = columnIndex, columnCount });
    }

    public string MergeTableCells(string sessionId, int tableIndex, int rowIndex, int startColumnIndex, int endColumnIndex)
    {
        if (tableIndex < 1) throw new InvalidOperationException("tableIndex must be >= 1.");
        if (rowIndex < 1) throw new InvalidOperationException("rowIndex must be >= 1.");
        if (startColumnIndex < 1) throw new InvalidOperationException("startColumnIndex must be >= 1.");
        if (endColumnIndex < startColumnIndex) throw new InvalidOperationException("endColumnIndex must be >= startColumnIndex.");

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var mergedCellCount = 0;
        sessionManager.ExecuteWriteOperation(session, "word_merge_table_cells", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var table = GetWordTableByIndex(body, tableIndex);
            var rows = table.Elements<W.TableRow>().ToList();

            if (rowIndex > rows.Count)
            {
                throw new InvalidOperationException($"rowIndex {rowIndex} is out of range. Table {tableIndex} has {rows.Count} rows.");
            }

            var row = rows[rowIndex - 1];
            var cells = row.Elements<W.TableCell>().ToList();

            if (endColumnIndex > cells.Count)
            {
                throw new InvalidOperationException($"endColumnIndex {endColumnIndex} is out of range. Row has {cells.Count} columns.");
            }

            if (startColumnIndex == endColumnIndex)
            {
                throw new InvalidOperationException("startColumnIndex and endColumnIndex must differ; nothing to merge.");
            }

            var span = endColumnIndex - startColumnIndex + 1;

            // Set GridSpan on the first cell
            var firstCell = cells[startColumnIndex - 1];
            var tcPr = firstCell.TableCellProperties ?? new W.TableCellProperties();
            if (firstCell.TableCellProperties == null)
            {
                firstCell.InsertAt(tcPr, 0);
            }

            var existingGridSpan = tcPr.GetFirstChild<W.GridSpan>();
            if (existingGridSpan != null)
            {
                existingGridSpan.Val = span;
            }
            else
            {
                tcPr.Append(new W.GridSpan { Val = span });
            }

            // Remove the spanned cells
            for (var ci = endColumnIndex - 1; ci >= startColumnIndex; ci--)
            {
                cells[ci].Remove();
                mergedCellCount++;
            }

            document.MainDocumentPart!.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_merge_table_cells", new { tableIndex, rowIndex, startColumnIndex, endColumnIndex, mergedCellCount });
    }

    public string DeleteParagraph(string sessionId, int paragraphIndex)
    {
        if (paragraphIndex < 1) throw new InvalidOperationException("paragraphIndex must be >= 1.");

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_delete_paragraph", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();

            if (paragraphIndex > paragraphs.Count)
            {
                throw new InvalidOperationException(BuildWordBodyParagraphRangeMessage(body, paragraphIndex, paragraphs.Count));
            }

            if (paragraphs.Count <= 1)
            {
                throw new InvalidOperationException("Cannot delete the only paragraph in the document. Word requires at least one paragraph.");
            }

            paragraphs[paragraphIndex - 1].Remove();
            document.MainDocumentPart!.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_delete_paragraph", new { removedIndex = paragraphIndex });
    }

    public string DeleteStyle(string sessionId, string styleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var deletedStyleId = string.Empty;
        var orphanedReferences = new List<object>();

        sessionManager.ExecuteWriteOperation(session, "word_delete_style", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            var stylePart = mainPart.StyleDefinitionsPart ?? throw new InvalidOperationException("Style definitions part is missing.");
            var styles = stylePart.Styles ?? throw new InvalidOperationException("Styles are missing.");

            // Resolve style by id or name
            var style = styles.Elements<W.Style>().FirstOrDefault(s =>
                string.Equals(s.StyleId?.Value, styleName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(s.StyleName?.Val?.Value, styleName, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException($"Style '{styleName}' was not found.");

            // Protect built-in styles
            var styleId = style.StyleId?.Value ?? styleName;
            var isBuiltIn = WordBuiltInStyles.Any(b =>
                string.Equals(b.StyleId, styleId, StringComparison.OrdinalIgnoreCase)
                || string.Equals(b.Name, styleName, StringComparison.OrdinalIgnoreCase));

            if (isBuiltIn)
            {
                throw new InvalidOperationException($"Style '{styleName}' is a built-in style and cannot be deleted.");
            }

            deletedStyleId = styleId;

            // Remove style definition
            style.Remove();
            stylePart.Styles.Save();

            // Scan for orphaned references
            var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var bodyParagraphs = body.Elements<W.Paragraph>().ToList();

            for (var pi = 0; pi < bodyParagraphs.Count; pi++)
            {
                var para = bodyParagraphs[pi];
                var paraStyleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (string.Equals(paraStyleId, deletedStyleId, StringComparison.OrdinalIgnoreCase))
                {
                    orphanedReferences.Add(new
                    {
                        type = "paragraph",
                        paragraphIndex = pi + 1,
                        preview = TrimPreview(para.InnerText)
                    });
                }

                var runs = para.Elements<W.Run>().ToList();
                for (var ri = 0; ri < runs.Count; ri++)
                {
                    var runStyleId = runs[ri].RunProperties?.RunStyle?.Val?.Value;
                    if (string.Equals(runStyleId, deletedStyleId, StringComparison.OrdinalIgnoreCase))
                    {
                        orphanedReferences.Add(new
                        {
                            type = "run",
                            paragraphIndex = pi + 1,
                            runIndex = ri + 1,
                            preview = TrimPreview(runs[ri].InnerText)
                        });
                    }
                }
            }

            mainPart.Document?.Save();
        });

        return SessionManager.BuildResult("word_delete_style", new
        {
            deletedStyle = deletedStyleId,
            orphanedReferenceCount = orphanedReferences.Count,
            orphanedReferences
        });
    }

    public string AddBulletedList(string sessionId, string lines, string bulletStyle = "disc")
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(lines);
        var items = lines.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(x => new WordListItem(x, 0, "bulleted", BulletStyle: bulletStyle))
            .ToArray();
        return AddStructuredList(sessionId, JsonSerializer.Serialize(items), "word_add_bulleted_list");
    }

    public string AddNumberedList(string sessionId, string lines, string numberStyle = "decimal-dot")
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(lines);
        var items = lines.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(x => new WordListItem(x, 0, "numbered", NumberStyle: numberStyle))
            .ToArray();
        return AddStructuredList(sessionId, JsonSerializer.Serialize(items), "word_add_numbered_list");
    }

    public string AddStructuredList(string sessionId, string itemsJson)
    {
        return AddStructuredList(sessionId, itemsJson, "word_add_structured_list");
    }

    public string InsertParagraphAfterText(string sessionId, string anchorText, string text, int occurrence = 1, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(anchorText);
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        if (occurrence < 1)
        {
            throw new InvalidOperationException("Occurrence must be >= 1.");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var insertedIndex = -1;
        sessionManager.ExecuteWriteOperation(session, "word_insert_paragraph_after_text", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            var match = paragraphs
                .Select((p, i) => new { Paragraph = p, Index = i + 1 })
                .Where(x => x.Paragraph.InnerText.Contains(anchorText, comparison))
                .Skip(occurrence - 1)
                .FirstOrDefault()
                ?? throw new InvalidOperationException($"Anchor text '{anchorText}' occurrence {occurrence} was not found.");

            var newParagraph = new W.Paragraph(new W.Run(new W.Text(text)));
            ApplyWordParagraphSpacing(newParagraph, DefaultParagraphBeforePt, DefaultParagraphAfterPt, DefaultParagraphLineSpacing);
            match.Paragraph.InsertAfterSelf(newParagraph);
            insertedIndex = match.Index + 1;

            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_insert_paragraph_after_text", new { insertedIndex, occurrence });
    }

    public string InsertTextAfterText(string sessionId, string anchorText, string text, int occurrence = 1, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(anchorText);
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        if (occurrence < 1)
        {
            throw new InvalidOperationException("Occurrence must be >= 1.");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var replaced = false;
        sessionManager.ExecuteWriteOperation(session, "word_insert_text_after_text", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var textNodes = document.MainDocumentPart?.Document?.Descendants<W.Text>().ToList()
                ?? throw new InvalidOperationException("Word text nodes are missing.");

            var remaining = occurrence;
            foreach (var node in textNodes)
            {
                if (string.IsNullOrEmpty(node.Text))
                {
                    continue;
                }

                var index = node.Text.IndexOf(anchorText, comparison);
                if (index < 0)
                {
                    continue;
                }

                remaining--;
                if (remaining > 0)
                {
                    continue;
                }

                var insertPos = index + anchorText.Length;
                node.Text = node.Text[..insertPos] + text + node.Text[insertPos..];
                replaced = true;
                break;
            }

            if (!replaced)
            {
                throw new InvalidOperationException($"Anchor text '{anchorText}' occurrence {occurrence} was not found.");
            }

            document.MainDocumentPart?.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_insert_text_after_text", new { occurrence }, replaced);
    }

    public string ListStyles(string sessionId)
    {
        var session = sessionManager.GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Word)
        {
            throw new InvalidOperationException("Session document type is not Word.");
        }

        using var document = WordprocessingDocument.Open(session.FilePath, false);
        var styles = (document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<W.Style>() ?? []).ToList();
        var builtInById = WordBuiltInStyles.ToDictionary<WordBuiltInStyleDefinition, string>(s => s.StyleId, StringComparer.OrdinalIgnoreCase);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var payload = new List<object>();

        foreach (var definition in WordBuiltInStyles)
        {
            var existing = styles.FirstOrDefault(s => string.Equals(s.StyleId?.Value, definition.StyleId, StringComparison.OrdinalIgnoreCase));
            payload.Add(new
            {
                styleId = existing?.StyleId?.Value ?? definition.StyleId,
                name = existing?.StyleName?.Val?.Value ?? definition.Name,
                type = existing is not null ? ResolveWordStyleType(existing) : definition.Type,
                isDefault = existing?.Default?.Value ?? definition.IsDefault,
                isBuiltIn = true
            });

            seen.Add(definition.StyleId);
        }

        foreach (var style in styles)
        {
            var styleId = style.StyleId?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(styleId) || seen.Contains(styleId))
            {
                continue;
            }

            payload.Add(new
            {
                styleId,
                name = style.StyleName?.Val?.Value ?? string.Empty,
                type = ResolveWordStyleType(style),
                isDefault = style.Default?.Value ?? false,
                isBuiltIn = builtInById.ContainsKey(styleId)
            });
            seen.Add(styleId);
        }

        return JsonSerializer.Serialize(new { count = payload.Count, styles = payload });
    }

    public string ApplyStyleByName(string sessionId, int paragraphIndex, string styleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_apply_style_by_name", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            EnsureWordStyleInfrastructure(mainPart);
            var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();
            if (paragraphIndex < 1 || paragraphIndex > paragraphs.Count)
            {
                throw new InvalidOperationException(BuildWordBodyParagraphRangeMessage(body, paragraphIndex, paragraphs.Count));
            }

            var styles = mainPart.StyleDefinitionsPart?.Styles?.Elements<W.Style>() ?? [];
            var style = styles.FirstOrDefault(s =>
                string.Equals(s.StyleId?.Value, styleName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(s.StyleName?.Val?.Value, styleName, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException($"Style '{styleName}' was not found in document.");

            if (style.Type?.Value != W.StyleValues.Paragraph)
            {
                throw new InvalidOperationException($"Style '{styleName}' is not a paragraph style.");
            }

            var paragraph = paragraphs[paragraphIndex - 1];
            var paragraphProperties = paragraph.ParagraphProperties ??= new W.ParagraphProperties();
            paragraphProperties.ParagraphStyleId = new W.ParagraphStyleId { Val = style.StyleId?.Value };
            mainPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_apply_style_by_name", new { paragraphIndex, styleName });
    }

    public string CreateOrUpdateStyle(string sessionId, string styleName, string styleJson)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        ArgumentException.ThrowIfNullOrWhiteSpace(styleJson);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_create_or_update_style", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            EnsureWordStyleInfrastructure(mainPart);
            var stylePart = mainPart.StyleDefinitionsPart ?? throw new InvalidOperationException("Style definitions part is missing.");
            var styles = stylePart.Styles ?? throw new InvalidOperationException("Styles are missing.");

            var options = JsonNode.Parse(styleJson) as JsonObject ?? throw new InvalidOperationException("Invalid styleJson object.");
            var requestedStyleType = ParseWordStyleType(options["type"]?.GetValue<string>() ?? "paragraph");

            var styleId = SanitizeStyleId(styleName);
            var style = styles.Elements<W.Style>().FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
            if (style is null)
            {
                style = new W.Style { Type = requestedStyleType, StyleId = styleId, CustomStyle = true };
                style.Append(new W.StyleName { Val = styleName });
                if (requestedStyleType == W.StyleValues.Paragraph)
                {
                    style.Append(new W.BasedOn { Val = "Normal" });
                    style.Append(new W.NextParagraphStyle { Val = "Normal" });
                }

                style.Append(new W.UIPriority { Val = 99 });
                style.Append(new W.UnhideWhenUsed());
                style.Append(new W.PrimaryStyle());
                styles.Append(style);
            }
            else if (style.Type?.Value is not null && style.Type.Value != requestedStyleType)
            {
                throw new InvalidOperationException($"Style '{styleName}' already exists with type '{ResolveWordStyleType(style)}'.");
            }

            ApplyStyleOptions(style, options, requestedStyleType);
            stylePart.Styles.Save();
            mainPart.Document?.Save();
        });

        var parsedOptions = JsonNode.Parse(styleJson) as JsonObject;
        var type = parsedOptions is null ? "paragraph" : ResolveWordStyleType(ParseWordStyleType(parsedOptions["type"]?.GetValue<string>() ?? "paragraph"));
        return SessionManager.BuildMutationResult("word_create_or_update_style", new { styleName, styleId = SanitizeStyleId(styleName), type });
    }

    public string ApplyCharacterStyleToText(string sessionId, string anchorText, string styleName, int occurrence = 1, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(anchorText);
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        if (occurrence <= 0)
        {
            throw new InvalidOperationException("Occurrence must be greater than zero.");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_apply_character_style_to_text", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            EnsureWordStyleInfrastructure(mainPart);

            var style = ResolveWordStyle(mainPart, styleName)
                ?? throw new InvalidOperationException($"Style '{styleName}' was not found in document.");
            if (style.Type?.Value != W.StyleValues.Character)
            {
                throw new InvalidOperationException($"Style '{styleName}' is not a character style.");
            }

            var paragraphs = mainPart.Document?.Body?.Descendants<W.Paragraph>() ?? throw new InvalidOperationException("Word document paragraphs are missing.");
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var remainingOccurrence = occurrence;

            foreach (var paragraph in paragraphs)
            {
                var segments = BuildParagraphRunSegments(paragraph).ToList();
                if (segments.Count == 0)
                {
                    continue;
                }

                var paragraphText = string.Concat(segments.Select(s => s.Text));
                var matches = FindWholeWordMatches(paragraphText, anchorText, comparison).ToList();
                if (matches.Count == 0)
                {
                    continue;
                }

                if (remainingOccurrence > matches.Count)
                {
                    remainingOccurrence -= matches.Count;
                    continue;
                }

                var match = matches[remainingOccurrence - 1];
                var matchStart = match.Start;
                var matchEnd = matchStart + match.Length;

                foreach (var segment in segments)
                {
                    var segmentEnd = segment.Start + segment.Length;
                    var overlapStart = Math.Max(matchStart, segment.Start);
                    var overlapEnd = Math.Min(matchEnd, segmentEnd);
                    if (overlapEnd <= overlapStart)
                    {
                        continue;
                    }

                    SplitRunAndApplyCharacterStyle(
                        segment.Run,
                        segment.Text,
                        overlapStart - segment.Start,
                        overlapEnd - overlapStart,
                        style.StyleId?.Value ?? string.Empty);
                }

                mainPart.Document?.Save();
                return;
            }

            throw new InvalidOperationException($"Anchor text '{anchorText}' occurrence {occurrence} was not found.");
        });

        return SessionManager.BuildMutationResult("word_apply_character_style_to_text", new { styleName, occurrence });
    }

    public string ApplyCharacterStyleToAll(string sessionId, string queriesJson, string styleName, bool matchCase = false, bool wholeWord = true)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(queriesJson);
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);

        var queries = JsonSerializer.Deserialize<List<string>>(queriesJson, new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("queriesJson must be a JSON array of strings.");
        if (queries.Count == 0)
        {
            throw new InvalidOperationException("At least one query is required.");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var queryResults = queries.Select(q => new { query = q, matchCount = 0, styledCount = 0 }).ToList();
        var totalStyled = 0;

        sessionManager.ExecuteWriteOperation(session, "word_apply_character_style_to_all", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            EnsureWordStyleInfrastructure(mainPart);

            var style = ResolveWordStyle(mainPart, styleName)
                ?? throw new InvalidOperationException($"Style '{styleName}' was not found in document.");
            if (style.Type?.Value != W.StyleValues.Character)
            {
                throw new InvalidOperationException($"Style '{styleName}' is not a character style.");
            }

            var styleId = style.StyleId?.Value ?? string.Empty;
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var paragraphs = mainPart.Document?.Body?.Descendants<W.Paragraph>().ToList()
                ?? throw new InvalidOperationException("Word document paragraphs are missing.");

            for (var qi = 0; qi < queries.Count; qi++)
            {
                var query = queries[qi];
                var matchCount = 0;
                var styledCount = 0;

                // Process paragraphs in reverse to avoid index shifting issues when splitting runs
                foreach (var paragraph in paragraphs)
                {
                    var segments = BuildParagraphRunSegments(paragraph).ToList();
                    if (segments.Count == 0) continue;

                    var paragraphText = string.Concat(segments.Select(s => s.Text));
                    var matches = wholeWord
                        ? FindWholeWordMatches(paragraphText, query, comparison)
                        : FindAllMatches(paragraphText, query, comparison);

                    if (matches.Count == 0) continue;
                    matchCount += matches.Count;

                    // Apply in reverse order so earlier offsets are unaffected by splits
                    foreach (var match in Enumerable.Reverse(matches))
                    {
                        var matchStart = match.Start;
                        var matchEnd = matchStart + match.Length;

                        // Rebuild segments after each apply since run structure changes
                        var currentSegments = BuildParagraphRunSegments(paragraph).ToList();
                        var currentText = string.Concat(currentSegments.Select(s => s.Text));

                        // Verify the match still exists at this position
                        if (matchStart >= currentText.Length)
                        {
                            continue;
                        }

                        foreach (var segment in currentSegments)
                        {
                            var segmentEnd = segment.Start + segment.Length;
                            var overlapStart = Math.Max(matchStart, segment.Start);
                            var overlapEnd = Math.Min(matchEnd, segmentEnd);
                            if (overlapEnd <= overlapStart) continue;

                            SplitRunAndApplyCharacterStyle(
                                segment.Run, segment.Text,
                                overlapStart - segment.Start,
                                overlapEnd - overlapStart,
                                styleId);
                        }

                        styledCount++;
                        totalStyled++;
                    }
                }

                queryResults[qi] = new { query, matchCount, styledCount };
            }

            mainPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_apply_character_style_to_all",
            new { styleName, totalStyled, queryResults });
    }

    public string ApplyCharacterStyleByPattern(string sessionId, string pattern, string styleName, bool matchCase = true, int maxMatches = 5000)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pattern);
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        if (maxMatches < 1 || maxMatches > 50000)
        {
            throw new InvalidOperationException("maxMatches must be between 1 and 50000.");
        }

        // Validate regex before opening doc
        System.Text.RegularExpressions.Regex regex;
        try
        {
            var options = matchCase
                ? System.Text.RegularExpressions.RegexOptions.None
                : System.Text.RegularExpressions.RegexOptions.IgnoreCase;
            regex = new System.Text.RegularExpressions.Regex(pattern, options, TimeSpan.FromSeconds(5));
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Invalid regex pattern: {ex.Message}");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var totalStyled = 0;
        var totalMatches = 0;

        sessionManager.ExecuteWriteOperation(session, "word_apply_character_style_by_pattern", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            EnsureWordStyleInfrastructure(mainPart);

            var style = ResolveWordStyle(mainPart, styleName)
                ?? throw new InvalidOperationException($"Style '{styleName}' was not found in document.");
            if (style.Type?.Value != W.StyleValues.Character)
            {
                throw new InvalidOperationException($"Style '{styleName}' is not a character style.");
            }

            var styleId = style.StyleId?.Value ?? string.Empty;
            var paragraphs = mainPart.Document?.Body?.Descendants<W.Paragraph>().ToList()
                ?? throw new InvalidOperationException("Word document paragraphs are missing.");

            foreach (var paragraph in paragraphs)
            {
                if (totalStyled >= maxMatches) break;

                var segments = BuildParagraphRunSegments(paragraph).ToList();
                if (segments.Count == 0) continue;

                var paragraphText = string.Concat(segments.Select(s => s.Text));
                var matches = new List<(int Start, int Length)>();
                try
                {
                    foreach (System.Text.RegularExpressions.Match m in regex.Matches(paragraphText))
                    {
                        if (m.Length == 0) continue;
                        matches.Add((m.Index, m.Length));
                    }
                }
                catch (System.Text.RegularExpressions.RegexMatchTimeoutException)
                {
                    throw new InvalidOperationException($"Pattern matching timed out on a paragraph. Simplify the pattern.");
                }

                if (matches.Count == 0) continue;
                totalMatches += matches.Count;

                foreach (var match in Enumerable.Reverse(matches))
                {
                    if (totalStyled >= maxMatches) break;

                    var matchStart = match.Start;
                    var matchEnd = matchStart + match.Length;
                    var currentSegments = BuildParagraphRunSegments(paragraph).ToList();

                    foreach (var segment in currentSegments)
                    {
                        var segmentEnd = segment.Start + segment.Length;
                        var overlapStart = Math.Max(matchStart, segment.Start);
                        var overlapEnd = Math.Min(matchEnd, segmentEnd);
                        if (overlapEnd <= overlapStart) continue;

                        SplitRunAndApplyCharacterStyle(
                            segment.Run, segment.Text,
                            overlapStart - segment.Start,
                            overlapEnd - overlapStart,
                            styleId);
                    }

                    totalStyled++;
                }
            }

            mainPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_apply_character_style_by_pattern",
            new { pattern, styleName, totalMatches, totalStyled });
    }

    public string InsertAfterHeading(string sessionId, string headingText, string text, int occurrence = 1, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(headingText);
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        if (occurrence < 1) throw new InvalidOperationException("Occurrence must be >= 1.");

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var insertedIndex = -1;

        sessionManager.ExecuteWriteOperation(session, "word_insert_after_heading", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            var bodyChildren = body.Elements().ToList();
            var headingMatches = bodyChildren
                .OfType<W.Paragraph>()
                .Where(p => GetHeadingLevel(p).HasValue && p.InnerText.Contains(headingText, comparison))
                .ToList();

            if (headingMatches.Count < occurrence)
            {
                throw new InvalidOperationException($"Heading containing '{headingText}' occurrence {occurrence} was not found (found {headingMatches.Count}).");
            }

            var targetHeading = headingMatches[occurrence - 1];
            var newParagraph = new W.Paragraph(new W.Run(new W.Text(text)));
            ApplyWordParagraphSpacing(newParagraph, DefaultParagraphBeforePt, DefaultParagraphAfterPt, DefaultParagraphLineSpacing);
            targetHeading.InsertAfterSelf(newParagraph);

            var bodyParagraphs = body.Elements<W.Paragraph>().ToList();
            insertedIndex = bodyParagraphs.IndexOf(newParagraph) + 1;

            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_insert_after_heading", new { headingText, insertedIndex, occurrence });
    }

    public string ReplaceSection(string sessionId, string headingText, string replacementJson, int occurrence = 1, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(headingText);
        ArgumentException.ThrowIfNullOrWhiteSpace(replacementJson);
        if (occurrence < 1) throw new InvalidOperationException("Occurrence must be >= 1.");

        var replacementParagraphs = JsonSerializer.Deserialize<List<string>>(replacementJson, new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("replacementJson must be a JSON array of paragraph strings.");

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var removedCount = 0;
        var insertedCount = 0;

        sessionManager.ExecuteWriteOperation(session, "word_replace_section", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            var bodyChildren = body.Elements().ToList();
            var headingMatches = bodyChildren
                .OfType<W.Paragraph>()
                .Where(p => GetHeadingLevel(p).HasValue && p.InnerText.Contains(headingText, comparison))
                .ToList();

            if (headingMatches.Count < occurrence)
            {
                throw new InvalidOperationException($"Heading containing '{headingText}' occurrence {occurrence} was not found (found {headingMatches.Count}).");
            }

            var targetHeading = headingMatches[occurrence - 1];
            var targetLevel = GetHeadingLevel(targetHeading)!.Value;

            // Collect all body elements between this heading and the next heading of equal/higher level
            var toRemove = new List<OpenXmlElement>();
            var started = false;
            foreach (var child in bodyChildren)
            {
                if (ReferenceEquals(child, targetHeading))
                {
                    started = true;
                    continue;
                }

                if (!started) continue;

                if (child is W.Paragraph p && GetHeadingLevel(p) is int lvl && lvl <= targetLevel)
                {
                    break;
                }

                toRemove.Add(child);
            }

            // Insert new paragraphs after the heading, then remove old ones
            var anchor = (OpenXmlElement)targetHeading;
            foreach (var paraText in replacementParagraphs)
            {
                var newPara = new W.Paragraph(new W.Run(new W.Text(paraText)));
                ApplyWordParagraphSpacing(newPara, DefaultParagraphBeforePt, DefaultParagraphAfterPt, DefaultParagraphLineSpacing);
                anchor.InsertAfterSelf(newPara);
                anchor = newPara;
                insertedCount++;
            }

            foreach (var el in toRemove)
            {
                el.Remove();
                removedCount++;
            }

            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_replace_section",
            new { headingText, occurrence, removedCount, insertedCount });
    }

    public string InsertTableOfContents(string sessionId, int paragraphIndex, int minLevel = 1, int maxLevel = 3)
    {
        if (minLevel < 1 || minLevel > 9) throw new InvalidOperationException("minLevel must be between 1 and 9.");
        if (maxLevel < 1 || maxLevel > 9) throw new InvalidOperationException("maxLevel must be between 1 and 9.");
        if (maxLevel < minLevel) throw new InvalidOperationException("maxLevel must be >= minLevel.");

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_insert_table_of_contents", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();

            if (paragraphIndex < 1 || paragraphIndex > paragraphs.Count + 1)
            {
                throw new InvalidOperationException($"Paragraph index {paragraphIndex} is out of range. Valid range is 1..{paragraphs.Count + 1}.");
            }

            // Build TOC field code: TOC \o "1-3" \h \z \u
            var levelRange = $"{minLevel}-{maxLevel}";
            var fieldCode = $" TOC \\o \"{levelRange}\" \\h \\z \\u ";

            // A TOC is inserted as a paragraph containing a complex field: fldChar begin → instrText → fldChar separate → fldChar end
            var tocParagraph = new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = "TOC1" }),
                new W.Run(new W.RunProperties(new W.RunStyle { Val = "TOCHyperlink" }),
                    new W.FieldChar { FieldCharType = W.FieldCharValues.Begin, Dirty = true }),
                new W.Run(new W.RunProperties(new W.RunStyle { Val = "TOCHyperlink" }),
                    new W.FieldCode(fieldCode) { Space = SpaceProcessingModeValues.Preserve }),
                new W.Run(new W.RunProperties(new W.RunStyle { Val = "TOCHyperlink" }),
                    new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
                new W.Run(new W.Text("(Right-click to update field)") { Space = SpaceProcessingModeValues.Preserve }),
                new W.Run(new W.RunProperties(new W.RunStyle { Val = "TOCHyperlink" }),
                    new W.FieldChar { FieldCharType = W.FieldCharValues.End }));

            if (paragraphIndex > paragraphs.Count)
            {
                body.Append(tocParagraph);
            }
            else
            {
                paragraphs[paragraphIndex - 1].InsertBeforeSelf(tocParagraph);
            }

            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_insert_table_of_contents", new { paragraphIndex, minLevel, maxLevel });
    }

    public string InsertPageBreakAfter(string sessionId, int paragraphIndex)
    {
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_insert_page_break_after", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();
            if (paragraphIndex < 1 || paragraphIndex > paragraphs.Count)
            {
                throw new InvalidOperationException(BuildWordBodyParagraphRangeMessage(body, paragraphIndex, paragraphs.Count));
            }

            var breakParagraph = new W.Paragraph(
                new W.Run(new W.Break { Type = W.BreakValues.Page }));
            paragraphs[paragraphIndex - 1].InsertAfterSelf(breakParagraph);
            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_insert_page_break_after", new { paragraphIndex });
    }

    public string SetHeader(string sessionId, string text, int sectionIndex = 1)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_set_header", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");

            var headerPart = mainPart.HeaderParts.FirstOrDefault() ?? mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(BuildHeaderFooterParagraph(text));
            headerPart.Header.Save();

            var sectionProps = mainPart.Document?.Body?.Elements<W.SectionProperties>().FirstOrDefault()
                ?? mainPart.Document?.Body?.AppendChild(new W.SectionProperties());
            if (sectionProps is null) throw new InvalidOperationException("Could not access section properties.");

            // Remove existing header references, then add new one
            sectionProps.Elements<W.HeaderReference>().ToList().ForEach(r => r.Remove());
            sectionProps.Append(new W.HeaderReference
            {
                Type = W.HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(headerPart)
            });

            mainPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_set_header", new { sectionIndex });
    }

    public string SetFooter(string sessionId, string text, int sectionIndex = 1)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_set_footer", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");

            var footerPart = mainPart.FooterParts.FirstOrDefault() ?? mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(BuildHeaderFooterParagraph(text));
            footerPart.Footer.Save();

            var sectionProps = mainPart.Document?.Body?.Elements<W.SectionProperties>().FirstOrDefault()
                ?? mainPart.Document?.Body?.AppendChild(new W.SectionProperties());
            if (sectionProps is null) throw new InvalidOperationException("Could not access section properties.");

            sectionProps.Elements<W.FooterReference>().ToList().ForEach(r => r.Remove());
            sectionProps.Append(new W.FooterReference
            {
                Type = W.HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(footerPart)
            });

            mainPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_set_footer", new { sectionIndex });
    }

    /// <summary>
    /// Builds a paragraph for a header or footer, substituting {PAGE} and {NUMPAGES} tokens with Word field codes.
    /// </summary>
    private static W.Paragraph BuildHeaderFooterParagraph(string text)
    {
        var para = new W.Paragraph();
        var pPr = new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Header" });
        para.Append(pPr);

        // Split text on known tokens and build runs + field codes
        var tokens = SplitOnTokens(text, ["{PAGE}", "{NUMPAGES}"]);
        foreach (var token in tokens)
        {
            switch (token)
            {
                case "{PAGE}":
                    para.Append(BuildSimpleFieldRun("PAGE"));
                    break;
                case "{NUMPAGES}":
                    para.Append(BuildSimpleFieldRun("NUMPAGES"));
                    break;
                default:
                    if (!string.IsNullOrEmpty(token))
                    {
                        para.Append(new W.Run(new W.Text(token) { Space = SpaceProcessingModeValues.Preserve }));
                    }
                    break;
            }
        }

        return para;
    }

    private static IEnumerable<string> SplitOnTokens(string text, string[] tokens)
    {
        var parts = new List<string>();
        var remaining = text;
        while (remaining.Length > 0)
        {
            var earliest = tokens
                .Select(t => new { token = t, index = remaining.IndexOf(t, StringComparison.Ordinal) })
                .Where(x => x.index >= 0)
                .OrderBy(x => x.index)
                .FirstOrDefault();

            if (earliest is null)
            {
                parts.Add(remaining);
                break;
            }

            if (earliest.index > 0)
            {
                parts.Add(remaining[..earliest.index]);
            }

            parts.Add(earliest.token);
            remaining = remaining[(earliest.index + earliest.token.Length)..];
        }

        return parts;
    }

    private static W.Run[] BuildSimpleFieldRun(string fieldName)
    {
        return
        [
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode($" {fieldName} ") { Space = SpaceProcessingModeValues.Preserve }),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            new W.Run(new W.Text("0") { Space = SpaceProcessingModeValues.Preserve }),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End })
        ];
    }

    public string UpdateStyle(string sessionId, string styleName, string styleJson)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        ArgumentException.ThrowIfNullOrWhiteSpace(styleJson);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_update_style", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            EnsureWordStyleInfrastructure(mainPart);
            var stylePart = mainPart.StyleDefinitionsPart ?? throw new InvalidOperationException("Style definitions part is missing.");
            var styles = stylePart.Styles ?? throw new InvalidOperationException("Styles are missing.");

            // Resolve by styleId or name (includes built-ins)
            var style = styles.Elements<W.Style>().FirstOrDefault(s =>
                string.Equals(s.StyleId?.Value, styleName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(s.StyleName?.Val?.Value, styleName, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException($"Style '{styleName}' was not found. Ensure it is seeded (apply_style_preset or create_or_update_style first).");

            var options = JsonNode.Parse(styleJson) as JsonObject
                ?? throw new InvalidOperationException("Invalid styleJson object.");

            // Infer type from existing style, not from options (update never changes type)
            var styleType = style.Type?.Value ?? W.StyleValues.Paragraph;
            ApplyStyleOptions(style, options, styleType);
            stylePart.Styles.Save();
            mainPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_update_style", new { styleName });
    }

    public string AppendMarkdown(string sessionId, string markdown)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(markdown);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        var addedCount = 0;

        sessionManager.ExecuteWriteOperation(session, "word_append_markdown", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            EnsureWordStyleInfrastructure(mainPart);
            var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");

            var elements = ParseMarkdownToElements(markdown, document);
            foreach (var el in elements)
            {
                body.Append(el);
                addedCount++;
            }

            mainPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult("word_append_markdown", new { addedCount });
    }

    public string ValidateDocument(string sessionId)
    {
        var session = sessionManager.GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Word)
        {
            throw new InvalidOperationException("Session document type is not Word.");
        }

        using var document = WordprocessingDocument.Open(session.FilePath, false);
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
        var styles = mainPart.StyleDefinitionsPart?.Styles?.Elements<W.Style>().ToList() ?? [];

        var issues = new List<object>();

        // 1. Collect all headings and check for skipped levels
        var bodyParagraphs = body.Elements<W.Paragraph>().ToList();
        var headings = bodyParagraphs
            .Select((p, i) => new { Paragraph = p, BodyIndex = i + 1, Level = GetHeadingLevel(p) })
            .Where(x => x.Level.HasValue)
            .ToList();

        var previousLevel = 0;
        foreach (var h in headings)
        {
            if (h.Level!.Value > previousLevel + 1)
            {
                issues.Add(new
                {
                    code = "SkippedHeadingLevel",
                    severity = "warning",
                    location = $"paragraph {h.BodyIndex}",
                    message = $"Heading level {h.Level} follows level {previousLevel} — level {previousLevel + 1} was skipped.",
                    suggestion = $"Add an intermediate Heading {previousLevel + 1} or promote this heading."
                });
            }
            previousLevel = h.Level.Value;
        }

        // 2. List all headings for reference
        var headingList = headings.Select(h => new { level = h.Level, text = TrimPreview(h.Paragraph.InnerText), bodyIndex = h.BodyIndex }).ToArray();

        // 3. Empty table cells
        var tables = body.Elements<W.Table>().ToList();
        for (var ti = 0; ti < tables.Count; ti++)
        {
            var rows = tables[ti].Elements<W.TableRow>().ToList();
            var colCounts = rows.Select(r => r.Elements<W.TableCell>().Count()).ToList();
            var expectedCols = colCounts.Count > 0 ? colCounts[0] : 0;

            for (var ri = 0; ri < rows.Count; ri++)
            {
                var cells = rows[ri].Elements<W.TableCell>().ToList();
                // Inconsistent column count
                if (cells.Count != expectedCols)
                {
                    issues.Add(new
                    {
                        code = "InconsistentColumnCount",
                        severity = "warning",
                        location = $"table {ti + 1}, row {ri + 1}",
                        message = $"Row has {cells.Count} columns but table header row has {expectedCols}.",
                        suggestion = "Ensure all rows have the same number of columns."
                    });
                }

                for (var ci = 0; ci < cells.Count; ci++)
                {
                    if (string.IsNullOrWhiteSpace(cells[ci].InnerText))
                    {
                        issues.Add(new
                        {
                            code = "EmptyTableCell",
                            severity = "info",
                            location = $"table {ti + 1}, row {ri + 1}, column {ci + 1}",
                            message = "Table cell is empty.",
                            suggestion = "Fill in the cell or use a placeholder."
                        });
                    }
                }
            }
        }

        // 4. Direct formatting (paragraphs with explicit run font/size/color not from a style)
        var directFormatCount = 0;
        foreach (var para in bodyParagraphs)
        {
            foreach (var run in para.Elements<W.Run>())
            {
                var rPr = run.RunProperties;
                if (rPr is null) continue;
                if (rPr.RunStyle is not null) continue; // has a character style — fine
                if (rPr.RunFonts is not null || rPr.FontSize is not null || rPr.Color is not null)
                {
                    directFormatCount++;
                    break; // count once per paragraph
                }
            }
        }

        if (directFormatCount > 0)
        {
            issues.Add(new
            {
                code = "DirectFormatting",
                severity = "info",
                location = "document",
                message = $"{directFormatCount} paragraph(s) contain runs with direct font/size/color formatting not from a character style.",
                suggestion = "Use character styles for consistent formatting."
            });
        }

        // 5. Unused custom styles
        var usedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var para in body.Descendants<W.Paragraph>())
        {
            var sid = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(sid)) usedStyleIds.Add(sid);
            foreach (var run in para.Elements<W.Run>())
            {
                var rsid = run.RunProperties?.RunStyle?.Val?.Value;
                if (!string.IsNullOrEmpty(rsid)) usedStyleIds.Add(rsid);
            }
        }

        var builtInIds = new HashSet<string>(WordBuiltInStyles.Select(s => s.StyleId), StringComparer.OrdinalIgnoreCase);
        var unusedCustomStyles = styles
            .Where(s => s.CustomStyle?.Value == true
                && !builtInIds.Contains(s.StyleId?.Value ?? string.Empty)
                && !usedStyleIds.Contains(s.StyleId?.Value ?? string.Empty))
            .Select(s => s.StyleId?.Value ?? string.Empty)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToArray();

        if (unusedCustomStyles.Length > 0)
        {
            issues.Add(new
            {
                code = "UnusedCustomStyle",
                severity = "info",
                location = "document",
                message = $"Custom style(s) defined but not used: {string.Join(", ", unusedCustomStyles)}.",
                suggestion = "Remove unused styles or apply them to appropriate content."
            });
        }

        return JsonSerializer.Serialize(new
        {
            ok = true,
            issueCount = issues.Count,
            headings = headingList,
            issues
        });
    }

    private static List<OpenXmlElement> ParseMarkdownToElements(string markdown, WordprocessingDocument document)
    {
        var elements = new List<OpenXmlElement>();
        var lines = markdown.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
        var i = 0;

        while (i < lines.Length)
        {
            var line = lines[i];

            // Fenced code block
            if (line.TrimStart().StartsWith("```", StringComparison.Ordinal))
            {
                i++;
                var codeLines = new List<string>();
                while (i < lines.Length && !lines[i].TrimStart().StartsWith("```", StringComparison.Ordinal))
                {
                    codeLines.Add(lines[i]);
                    i++;
                }
                i++; // skip closing ```
                foreach (var codeLine in codeLines)
                {
                    var p = new W.Paragraph(new W.Run(new W.Text(codeLine) { Space = SpaceProcessingModeValues.Preserve }));
                    elements.Add(p);
                }
                continue;
            }

            // Heading
            var headingMatch = System.Text.RegularExpressions.Regex.Match(line, @"^(#{1,9})\s+(.+)$");
            if (headingMatch.Success)
            {
                var level = headingMatch.Groups[1].Value.Length;
                var text = headingMatch.Groups[2].Value.Trim();
                var headingParagraph = new W.Paragraph(
                    new W.ParagraphProperties(new W.ParagraphStyleId { Val = $"Heading{level}" }),
                    new W.Run(new W.Text(text)));
                elements.Add(headingParagraph);
                i++;
                continue;
            }

            // Markdown table (strict: header row + separator row required)
            if (line.TrimStart().StartsWith("|", StringComparison.Ordinal))
            {
                var tableLines = new List<string>();
                while (i < lines.Length && lines[i].TrimStart().StartsWith("|", StringComparison.Ordinal))
                {
                    tableLines.Add(lines[i]);
                    i++;
                }
                var table = BuildMarkdownTable(tableLines);
                if (table is not null)
                {
                    elements.Add(table);
                    continue;
                }

                // Not a valid markdown table; preserve source as normal text paragraphs.
                foreach (var tableLikeLine in tableLines)
                {
                    var fallbackParagraph = new W.Paragraph();
                    AppendInlineMarkdown(fallbackParagraph, tableLikeLine);
                    elements.Add(fallbackParagraph);
                }
                continue;
            }

            // Unordered list item
            var ulMatch = System.Text.RegularExpressions.Regex.Match(line, @"^(\s*)[-*+]\s+(.+)$");
            if (ulMatch.Success)
            {
                var indent = ulMatch.Groups[1].Value.Length / 2;
                var text = ulMatch.Groups[2].Value;
                var listItems = new List<WordListItem> { new WordListItem(text, indent, "bulleted") };
                // collect continuation items
                i++;
                while (i < lines.Length)
                {
                    var nextUl = System.Text.RegularExpressions.Regex.Match(lines[i], @"^(\s*)[-*+]\s+(.+)$");
                    var nextOl = System.Text.RegularExpressions.Regex.Match(lines[i], @"^(\s*)\d+\.\s+(.+)$");
                    if (nextUl.Success)
                    {
                        listItems.Add(new WordListItem(nextUl.Groups[2].Value, nextUl.Groups[1].Value.Length / 2, "bulleted"));
                        i++;
                    }
                    else if (nextOl.Success)
                    {
                        listItems.Add(new WordListItem(nextOl.Groups[2].Value, nextOl.Groups[1].Value.Length / 2, "numbered"));
                        i++;
                    }
                    else break;
                }
                var listDefs = EnsureStructuredListNumbering(document, listItems);
                foreach (var item in listItems)
                {
                    var kind = NormalizeListKind(item.Kind);
                    var numId = kind == "numbered" ? listDefs.NumberedNumberingId : listDefs.BulletedNumberingId;
                    var listParagraph = new W.Paragraph(
                        new W.ParagraphProperties(
                            new W.NumberingProperties(
                                new W.NumberingLevelReference { Val = item.Level },
                                new W.NumberingId { Val = numId })));
                    AppendInlineMarkdown(listParagraph, item.Text);
                    elements.Add(listParagraph);
                }
                continue;
            }

            // Ordered list item
            var olMatch = System.Text.RegularExpressions.Regex.Match(line, @"^(\s*)\d+\.\s+(.+)$");
            if (olMatch.Success)
            {
                var indent = olMatch.Groups[1].Value.Length / 2;
                var text = olMatch.Groups[2].Value;
                var listItems = new List<WordListItem> { new WordListItem(text, indent, "numbered") };
                i++;
                while (i < lines.Length)
                {
                    var nextOl = System.Text.RegularExpressions.Regex.Match(lines[i], @"^(\s*)\d+\.\s+(.+)$");
                    var nextUl = System.Text.RegularExpressions.Regex.Match(lines[i], @"^(\s*)[-*+]\s+(.+)$");
                    if (nextOl.Success) { listItems.Add(new WordListItem(nextOl.Groups[2].Value, nextOl.Groups[1].Value.Length / 2, "numbered")); i++; }
                    else if (nextUl.Success) { listItems.Add(new WordListItem(nextUl.Groups[2].Value, nextUl.Groups[1].Value.Length / 2, "bulleted")); i++; }
                    else break;
                }
                var listDefs = EnsureStructuredListNumbering(document, listItems);
                foreach (var item in listItems)
                {
                    var kind = NormalizeListKind(item.Kind);
                    var numId = kind == "numbered" ? listDefs.NumberedNumberingId : listDefs.BulletedNumberingId;
                    var listParagraph = new W.Paragraph(
                        new W.ParagraphProperties(
                            new W.NumberingProperties(
                                new W.NumberingLevelReference { Val = item.Level },
                                new W.NumberingId { Val = numId })));
                    AppendInlineMarkdown(listParagraph, item.Text);
                    elements.Add(listParagraph);
                }
                continue;
            }

            // Blank line — skip
            if (string.IsNullOrWhiteSpace(line))
            {
                i++;
                continue;
            }

            // Normal paragraph with inline markdown.
            // Join wrapped source lines until a markdown block boundary or blank line.
            var paragraphLines = new List<string>();
            while (i < lines.Length)
            {
                var candidate = lines[i];
                if (string.IsNullOrWhiteSpace(candidate)
                    || System.Text.RegularExpressions.Regex.IsMatch(candidate, @"^(#{1,9})\s+(.+)$")
                    || System.Text.RegularExpressions.Regex.IsMatch(candidate, @"^(\s*)[-*+]\s+(.+)$")
                    || System.Text.RegularExpressions.Regex.IsMatch(candidate, @"^(\s*)\d+\.\s+(.+)$")
                    || candidate.TrimStart().StartsWith("|", StringComparison.Ordinal)
                    || candidate.TrimStart().StartsWith("```", StringComparison.Ordinal))
                {
                    break;
                }

                paragraphLines.Add(candidate.Trim());
                i++;
            }

            if (paragraphLines.Count == 0)
            {
                i++;
                continue;
            }

            var paragraph = new W.Paragraph();
            AppendInlineMarkdown(paragraph, string.Join(" ", paragraphLines));
            elements.Add(paragraph);
        }

        return elements;
    }

    private static W.Table? BuildMarkdownTable(List<string> tableLines)
    {
        if (tableLines.Count < 2)
        {
            return null;
        }

        var headerCells = SplitMarkdownTableRow(tableLines[0]);
        if (headerCells.Length < 2)
        {
            return null;
        }

        if (!IsMarkdownTableSeparatorRow(tableLines[1], headerCells.Length))
        {
            return null;
        }

        var rows = new List<string[]> { headerCells };
        for (var index = 2; index < tableLines.Count; index++)
        {
            var cells = SplitMarkdownTableRow(tableLines[index]);
            if (cells.Length != headerCells.Length)
            {
                return null;
            }

            rows.Add(cells);
        }

        var table = new W.Table();
        var colCount = rows[0].Length;
        for (var r = 0; r < rows.Count; r++)
        {
            var row = new W.TableRow();
            for (var c = 0; c < colCount; c++)
            {
                var cellText = c < rows[r].Length ? rows[r][c] : string.Empty;
                var cell = new W.TableCell();
                var para = new W.Paragraph();
                if (r == 0)
                {
                    // Header row — bold
                    var run = new W.Run(new W.RunProperties(new W.Bold()), new W.Text(cellText) { Space = SpaceProcessingModeValues.Preserve });
                    para.Append(run);
                }
                else
                {
                    AppendInlineMarkdown(para, cellText);
                }
                cell.Append(para);
                row.Append(cell);
            }
            table.Append(row);
        }

        return table;
    }

    private static string[] SplitMarkdownTableRow(string line)
    {
        return line.Trim().Trim('|').Split('|').Select(c => c.Trim()).ToArray();
    }

    private static bool IsMarkdownTableSeparatorRow(string line, int expectedColumns)
    {
        var cells = SplitMarkdownTableRow(line);
        if (cells.Length != expectedColumns)
        {
            return false;
        }

        return cells.All(cell => System.Text.RegularExpressions.Regex.IsMatch(cell, @"^:?-{3,}:?$"));
    }

    private static void AppendInlineMarkdown(W.Paragraph para, string text)
    {
        // Parse inline: **bold**, *italic*, `code`, and plain text
        var pattern = @"(\*\*(.+?)\*\*|\*(.+?)\*|`(.+?)`)";
        var lastIndex = 0;
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(text, pattern))
        {
            if (m.Index > lastIndex)
            {
                para.Append(new W.Run(new W.Text(text[lastIndex..m.Index]) { Space = SpaceProcessingModeValues.Preserve }));
            }

            if (m.Value.StartsWith("**", StringComparison.Ordinal))
            {
                para.Append(new W.Run(new W.RunProperties(new W.Bold()), new W.Text(m.Groups[2].Value) { Space = SpaceProcessingModeValues.Preserve }));
            }
            else if (m.Value.StartsWith("*", StringComparison.Ordinal))
            {
                para.Append(new W.Run(new W.RunProperties(new W.Italic()), new W.Text(m.Groups[3].Value) { Space = SpaceProcessingModeValues.Preserve }));
            }
            else if (m.Value.StartsWith("`", StringComparison.Ordinal))
            {
                // Code span — use IntenseReference style if available
                var run = new W.Run(
                    new W.RunProperties(new W.RunStyle { Val = "IntenseReference" }),
                    new W.Text(m.Groups[4].Value) { Space = SpaceProcessingModeValues.Preserve });
                para.Append(run);
            }

            lastIndex = m.Index + m.Length;
        }

        if (lastIndex < text.Length)
        {
            para.Append(new W.Run(new W.Text(text[lastIndex..]) { Space = SpaceProcessingModeValues.Preserve }));
        }

        // If nothing was appended (empty text), add empty run
        if (!para.Elements<W.Run>().Any())
        {
            para.Append(new W.Run(new W.Text(string.Empty)));
        }
    }

    public string ListParagraphRuns(string sessionId, int paragraphIndex)
    {
        var session = sessionManager.GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.Word)
        {
            throw new InvalidOperationException("Session document type is not Word.");
        }

        using var document = WordprocessingDocument.Open(session.FilePath, false);
        var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
        var paragraphs = body.Elements<W.Paragraph>().ToList();
        if (paragraphIndex < 1 || paragraphIndex > paragraphs.Count)
        {
            throw new InvalidOperationException(BuildWordBodyParagraphRangeMessage(body, paragraphIndex, paragraphs.Count));
        }

        var paragraph = paragraphs[paragraphIndex - 1];
        var styleNameById = BuildWordStyleNameMap(document.MainDocumentPart);
        var runs = paragraph.Elements<W.Run>().Select((run, index) =>
        {
            var runStyleId = run.RunProperties?.RunStyle?.Val?.Value ?? string.Empty;
            var fontSizeText = run.RunProperties?.FontSize?.Val?.Value;
            var fontSize = int.TryParse(fontSizeText, out var halfPoints) ? halfPoints / 2 : 0;
            return new
            {
                runIndex = index + 1,
                text = run.InnerText,
                styleId = runStyleId,
                styleName = runStyleId.Length > 0 && styleNameById.TryGetValue(runStyleId, out var resolvedName) ? resolvedName : string.Empty,
                fontName = run.RunProperties?.RunFonts?.Ascii?.Value ?? string.Empty,
                fontSize,
                bold = run.RunProperties?.Bold is not null,
                italic = run.RunProperties?.Italic is not null,
                colorHex = run.RunProperties?.Color?.Val?.Value ?? string.Empty
            };
        }).ToArray();

        return JsonSerializer.Serialize(new
        {
            paragraphIndex,
            paragraphText = paragraph.InnerText,
            runCount = runs.Length,
            runs
        });
    }

    // -------------------------------------------------------------------------
    // Internal helpers used by OfficeSessionService for cross-cutting features
    // -------------------------------------------------------------------------

    public void SetParagraphStyle(string sessionId, int paragraphIndex, TextStyle style)
    {
        ValidateTextStyle(style);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, "word_set_paragraph_style", () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var paragraphs = body.Elements<W.Paragraph>().ToList();

            if (paragraphIndex < 1 || paragraphIndex > paragraphs.Count)
            {
                throw new InvalidOperationException(BuildWordBodyParagraphRangeMessage(body, paragraphIndex, paragraphs.Count));
            }

            ApplyWordParagraphStyle(paragraphs[paragraphIndex - 1], style);
            document.MainDocumentPart.Document?.Save();
        });
    }

    public void ApplyStylePreset(string filePath)
    {
        using var document = WordprocessingDocument.Open(filePath, true);
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
        EnsureWordStyleInfrastructure(mainPart);
        mainPart.Document?.Save();
    }

    public void InitializeEmptyDocument(string filePath)
    {
        using var word = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
        var mainPart = word.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        EnsureWordStyleInfrastructure(mainPart);
        mainPart.Document.Save();
    }

    public object ListStructure(string filePath)
    {
        using var document = WordprocessingDocument.Open(filePath, false);
        var mainPart = document.MainDocumentPart;
        var body = mainPart?.Document?.Body;
        if (body is null)
        {
            return new { paragraphCount = 0, bodyParagraphCount = 0, tableCount = 0, elements = Array.Empty<object>(), tables = Array.Empty<object>() };
        }

        var listKindByNumberingId = BuildListKindMap(mainPart!);
        var paragraphs = GetWordParagraphReferences(body);
        var tableSummaries = BuildWordTableSummaries(body).ToArray();
        var elements = new List<object>();

        foreach (var element in body.Elements())
        {
            if (element is W.Paragraph paragraph)
            {
                var reference = paragraphs.FirstOrDefault(x => ReferenceEquals(x.Paragraph, paragraph));
                var headingLevel = GetHeadingLevel(paragraph);
                var numbering = paragraph.ParagraphProperties?.NumberingProperties;
                var numberingId = numbering?.NumberingId?.Val?.Value;
                var listLevel = numbering?.NumberingLevelReference?.Val?.Value;
                elements.Add(new
                {
                    type = headingLevel.HasValue ? "heading" : "paragraph",
                    bodyParagraphIndex = reference?.BodyParagraphIndex,
                    allParagraphIndex = reference?.AllParagraphIndex,
                    headingLevel,
                    isListItem = numberingId.HasValue,
                    listKind = numberingId.HasValue ? ResolveListKind(listKindByNumberingId, numberingId.Value) : string.Empty,
                    listLevel,
                    numberingId,
                    text = paragraph.InnerText,
                    preview = TrimPreview(paragraph.InnerText)
                });
            }
            else if (element is W.Table table)
            {
                var tableIndex = body.Elements<W.Table>().TakeWhile(t => !ReferenceEquals(t, table)).Count() + 1;
                var rows = table.Elements<W.TableRow>().Count();
                var columns = table.Elements<W.TableRow>().FirstOrDefault()?.Elements<W.TableCell>().Count() ?? 0;
                elements.Add(new
                {
                    type = "table",
                    tableIndex,
                    rows,
                    columns,
                    preview = TrimPreview(string.Join(" ", table.Descendants<W.Paragraph>().Select(p => p.InnerText).Where(t => !string.IsNullOrWhiteSpace(t))))
                });
            }
        }

        return new
        {
            paragraphCount = paragraphs.Count,
            bodyParagraphCount = paragraphs.Count(x => x.BodyParagraphIndex.HasValue),
            tableCount = tableSummaries.Length,
            elements,
            tables = tableSummaries
        };
    }

    public object FindText(string filePath, string query)
    {
        using var document = WordprocessingDocument.Open(filePath, false);
        var body = document.MainDocumentPart?.Document?.Body;
        if (body is null)
        {
            return new { matchCount = 0, matches = Array.Empty<object>() };
        }

        var references = GetWordParagraphReferences(body);
        var matches = references
            .Where(x => x.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
            .Select(x => new
            {
                index = x.AllParagraphIndex,
                bodyParagraphIndex = x.BodyParagraphIndex,
                inTable = x.InTable,
                tableIndex = x.TableIndex,
                rowIndex = x.RowIndex,
                columnIndex = x.ColumnIndex,
                addressableByParagraphTools = x.BodyParagraphIndex.HasValue,
                text = x.Text
            })
            .ToArray();
        return new { matchCount = matches.Length, matches };
    }

    // -------------------------------------------------------------------------
    // Style infrastructure (static, reusable)
    // -------------------------------------------------------------------------

    public static void EnsureWordStyleInfrastructure(MainDocumentPart mainPart)
    {
        var stylePart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylePart.Styles ??= new W.Styles();

        EnsureWordDocDefaults(stylePart.Styles);
        EnsureBuiltInWordStyles(stylePart.Styles);
        stylePart.Styles.Save();
    }

    // -------------------------------------------------------------------------
    // Private implementation details
    // -------------------------------------------------------------------------

    private string AddStructuredList(string sessionId, string itemsJson, string operationName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(itemsJson);
        var items = JsonSerializer.Deserialize<List<WordListItem>>(itemsJson, new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("Invalid itemsJson for structured list.");
        if (items.Count == 0)
        {
            throw new InvalidOperationException("At least one list item is required.");
        }

        if (items.Any(i => string.IsNullOrWhiteSpace(i.Text)))
        {
            throw new InvalidOperationException("List items must contain text.");
        }

        if (items.Any(i => i.Level < 0 || i.Level > 8))
        {
            throw new InvalidOperationException("List levels must be between 0 and 8.");
        }

        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.Word);
        sessionManager.ExecuteWriteOperation(session, operationName, () =>
        {
            using var document = WordprocessingDocument.Open(session.FilePath, true);
            var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var definitions = EnsureStructuredListNumbering(document, items);

            foreach (var item in items)
            {
                var normalizedKind = NormalizeListKind(item.Kind);
                var numberingId = normalizedKind == "numbered" ? definitions.NumberedNumberingId : definitions.BulletedNumberingId;
                var paragraph = new W.Paragraph(
                    new W.ParagraphProperties(
                        new W.NumberingProperties(
                            new W.NumberingLevelReference { Val = item.Level },
                            new W.NumberingId { Val = numberingId })),
                    new W.Run(new W.Text(item.Text)));
                ApplyWordParagraphSpacing(paragraph, DefaultListBeforePt, DefaultListAfterPt, DefaultListLineSpacing);
                body.Append(paragraph);
            }

            document.MainDocumentPart.Document?.Save();
        });

        return SessionManager.BuildMutationResult(operationName, new { itemCount = items.Count });
    }

    private static void EnsureWordDocDefaults(W.Styles styles)
    {
        if (styles.Elements<W.DocDefaults>().Any())
        {
            return;
        }

        styles.Append(new W.DocDefaults(
            new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
                new W.RunFonts { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi },
                new W.FontSize { Val = "22" },
                new W.FontSizeComplexScript { Val = "22" })),
            new W.ParagraphPropertiesDefault(new W.ParagraphPropertiesBaseStyle(
                new W.SpacingBetweenLines
                {
                    Before = "0",
                    After = "200",
                    Line = "276",
                    LineRule = W.LineSpacingRuleValues.Auto,
                    BeforeAutoSpacing = OnOffValue.FromBoolean(false),
                    AfterAutoSpacing = OnOffValue.FromBoolean(false)
                }))));
    }

    private static void EnsureBuiltInWordStyles(W.Styles styles)
    {
        foreach (var definition in WordBuiltInStyles)
        {
            var existing = styles.Elements<W.Style>().FirstOrDefault(s => string.Equals(s.StyleId?.Value, definition.StyleId, StringComparison.OrdinalIgnoreCase));
            if (existing is not null)
            {
                if (existing.StyleName is null)
                {
                    existing.StyleName = new W.StyleName { Val = definition.Name };
                }

                if (existing.StyleParagraphProperties is null || existing.StyleRunProperties is null)
                {
                    ApplyBuiltInWordStyleFormatting(existing, definition.StyleId, onlyIfMissing: true);
                }

                continue;
            }

            styles.Append(BuildBuiltInWordStyle(definition));
        }
    }

    private static W.Style BuildBuiltInWordStyle(WordBuiltInStyleDefinition definition)
    {
        var style = new W.Style
        {
            Type = ParseWordStyleType(definition.Type),
            StyleId = definition.StyleId,
            CustomStyle = false
        };

        style.Append(new W.StyleName { Val = definition.Name });
        if (definition.IsDefault)
        {
            style.Default = true;
        }

        if (!string.IsNullOrWhiteSpace(definition.BasedOn))
        {
            style.Append(new W.BasedOn { Val = definition.BasedOn });
        }

        if (!string.IsNullOrWhiteSpace(definition.NextStyle) && string.Equals(definition.Type, "paragraph", StringComparison.OrdinalIgnoreCase))
        {
            style.Append(new W.NextParagraphStyle { Val = definition.NextStyle });
        }

        if (definition.UiPriority is not null)
        {
            style.Append(new W.UIPriority { Val = definition.UiPriority.Value });
        }

        if (definition.UnhideWhenUsed)
        {
            style.Append(new W.UnhideWhenUsed());
        }

        if (definition.IsPrimary)
        {
            style.Append(new W.PrimaryStyle());
        }

        ApplyBuiltInWordStyleFormatting(style, definition.StyleId, onlyIfMissing: false);

        return style;
    }

    private static void ApplyBuiltInWordStyleFormatting(W.Style style, string styleId, bool onlyIfMissing)
    {
        if (!onlyIfMissing || style.StyleParagraphProperties is null)
        {
            style.StyleParagraphProperties = BuildBuiltInParagraphProperties(styleId);
        }

        if (!onlyIfMissing || style.StyleRunProperties is null)
        {
            style.StyleRunProperties = BuildBuiltInRunProperties(styleId);
        }
    }

    private static W.StyleParagraphProperties BuildBuiltInParagraphProperties(string styleId)
    {
        styleId = styleId.Trim();
        return styleId switch
        {
            "Heading1" => new W.StyleParagraphProperties(
                new W.KeepNext(),
                new W.KeepLines(),
                new W.SpacingBetweenLines { Before = "360", After = "80" },
                new W.OutlineLevel { Val = 0 }),
            "Heading2" => new W.StyleParagraphProperties(
                new W.KeepNext(),
                new W.KeepLines(),
                new W.SpacingBetweenLines { Before = "160", After = "80" },
                new W.OutlineLevel { Val = 1 }),
            "Heading3" => new W.StyleParagraphProperties(
                new W.KeepNext(),
                new W.KeepLines(),
                new W.SpacingBetweenLines { Before = "160", After = "80" },
                new W.OutlineLevel { Val = 2 }),
            "Heading4" => new W.StyleParagraphProperties(
                new W.KeepNext(),
                new W.KeepLines(),
                new W.SpacingBetweenLines { Before = "80", After = "40" },
                new W.OutlineLevel { Val = 3 }),
            "Heading5" => new W.StyleParagraphProperties(
                new W.KeepNext(),
                new W.KeepLines(),
                new W.SpacingBetweenLines { Before = "80", After = "40" },
                new W.OutlineLevel { Val = 4 }),
            "Heading6" => new W.StyleParagraphProperties(
                new W.KeepNext(),
                new W.KeepLines(),
                new W.SpacingBetweenLines { Before = "40", After = "0" },
                new W.OutlineLevel { Val = 5 }),
            "Heading7" => new W.StyleParagraphProperties(
                new W.KeepNext(),
                new W.KeepLines(),
                new W.SpacingBetweenLines { Before = "40", After = "0" },
                new W.OutlineLevel { Val = 6 }),
            "Heading8" => new W.StyleParagraphProperties(
                new W.KeepNext(),
                new W.KeepLines(),
                new W.SpacingBetweenLines { After = "0" },
                new W.OutlineLevel { Val = 7 }),
            "Heading9" => new W.StyleParagraphProperties(
                new W.KeepNext(),
                new W.KeepLines(),
                new W.SpacingBetweenLines { After = "0" },
                new W.OutlineLevel { Val = 8 }),
            "Title" => new W.StyleParagraphProperties(
                new W.SpacingBetweenLines { After = "80", Line = "240", LineRule = W.LineSpacingRuleValues.Auto },
                new W.ContextualSpacing { Val = OnOffValue.FromBoolean(true) }),
            "Subtitle" => new W.StyleParagraphProperties(),
            "Quote" => new W.StyleParagraphProperties(
                new W.SpacingBetweenLines { Before = "160" },
                new W.Justification { Val = W.JustificationValues.Center }),
            "IntenseQuote" => new W.StyleParagraphProperties(
                new W.ParagraphBorders(
                    new W.TopBorder { Val = W.BorderValues.Single, Color = "0F4761", Size = 4U, Space = 10U },
                    new W.BottomBorder { Val = W.BorderValues.Single, Color = "0F4761", Size = 4U, Space = 10U }),
                new W.SpacingBetweenLines { Before = "360", After = "360" },
                new W.Indentation { Left = "864", Right = "864" },
                new W.Justification { Val = W.JustificationValues.Center }),
            "Caption" => new W.StyleParagraphProperties(
                new W.SpacingBetweenLines { After = "200", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }),
            "NoSpacing" => new W.StyleParagraphProperties(
                new W.SpacingBetweenLines { After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }),
            "ListParagraph" => new W.StyleParagraphProperties(
                new W.Indentation { Left = "720" },
                new W.ContextualSpacing { Val = OnOffValue.FromBoolean(true) }),
            "Header" => new W.StyleParagraphProperties(
                new W.SpacingBetweenLines { After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }),
            "Footer" => new W.StyleParagraphProperties(
                new W.SpacingBetweenLines { After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto }),
            _ => new W.StyleParagraphProperties()
        };
    }

    private static W.StyleRunProperties BuildBuiltInRunProperties(string styleId)
    {
        styleId = styleId.Trim();
        return styleId switch
        {
            "Heading1" => new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                new W.Color { Val = "0F4761" },
                new W.FontSize { Val = "40" },
                new W.FontSizeComplexScript { Val = "40" }),
            "Heading2" => new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                new W.Color { Val = "0F4761" },
                new W.FontSize { Val = "32" },
                new W.FontSizeComplexScript { Val = "32" }),
            "Heading3" => new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                new W.Color { Val = "0F4761" },
                new W.FontSize { Val = "28" },
                new W.FontSizeComplexScript { Val = "28" }),
            "Heading4" => new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                new W.Italic(),
                new W.Color { Val = "0F4761" }),
            "Heading5" => new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                new W.Color { Val = "0F4761" }),
            "Heading6" => new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                new W.Italic(),
                new W.Color { Val = "595959" }),
            "Heading7" => new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                new W.Color { Val = "595959" }),
            "Heading8" => new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                new W.Italic(),
                new W.Color { Val = "272727" }),
            "Heading9" => new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                new W.Italic(),
                new W.Color { Val = "272727" }),
            "Title" => new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
                new W.FontSize { Val = "56" },
                new W.FontSizeComplexScript { Val = "56" }),
            "Subtitle" => new W.StyleRunProperties(
                new W.Color { Val = "595959" },
                new W.FontSize { Val = "28" },
                new W.FontSizeComplexScript { Val = "28" }),
            "Quote" => new W.StyleRunProperties(
                new W.Italic(),
                new W.Color { Val = "404040" }),
            "IntenseQuote" => new W.StyleRunProperties(
                new W.Italic(),
                new W.Color { Val = "0F4761" }),
            "SubtleEmphasis" => new W.StyleRunProperties(
                new W.Italic(),
                new W.Color { Val = "404040" }),
            "Emphasis" => new W.StyleRunProperties(
                new W.Italic()),
            "IntenseEmphasis" => new W.StyleRunProperties(
                new W.Italic(),
                new W.Color { Val = "0F4761" }),
            "Strong" => new W.StyleRunProperties(
                new W.Bold()),
            "SubtleReference" => new W.StyleRunProperties(
                new W.SmallCaps(),
                new W.Color { Val = "5A5A5A" }),
            "IntenseReference" => new W.StyleRunProperties(
                new W.Bold(),
                new W.SmallCaps(),
                new W.Color { Val = "0F4761" }),
            "BookTitle" => new W.StyleRunProperties(
                new W.Bold(),
                new W.Italic()),
            "Caption" => new W.StyleRunProperties(
                new W.Italic(),
                new W.Color { Val = "0E2841" },
                new W.FontSize { Val = "18" },
                new W.FontSizeComplexScript { Val = "18" }),
            _ => new W.StyleRunProperties()
        };
    }

    private static void ApplyWordParagraphStyle(W.Paragraph paragraph, TextStyle style)
    {
        foreach (var run in paragraph.Elements<W.Run>())
        {
            run.RunProperties ??= new W.RunProperties();
            ApplyWordRunProperties(run.RunProperties, style);
        }

        var paragraphProperties = paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        paragraphProperties.ParagraphMarkRunProperties ??= new W.ParagraphMarkRunProperties();
        ApplyWordParagraphMarkProperties(paragraphProperties.ParagraphMarkRunProperties, style);
    }

    private static void ApplyWordParagraphSpacing(W.Paragraph paragraph, int beforePt, int afterPt, double lineSpacing)
    {
        var paragraphProperties = paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        var spacing = paragraphProperties.SpacingBetweenLines ??= new W.SpacingBetweenLines();
        spacing.Before = (beforePt * 20).ToString();
        spacing.After = (afterPt * 20).ToString();
        spacing.Line = ((int)Math.Round(lineSpacing * 240)).ToString();
        spacing.LineRule = W.LineSpacingRuleValues.Auto;
    }

    private static void ApplyWordRunProperties(W.RunProperties runProperties, TextStyle style)
    {
        runProperties.RunFonts = new W.RunFonts { Ascii = style.FontName, HighAnsi = style.FontName, ComplexScript = style.FontName };
        runProperties.FontSize = new W.FontSize { Val = (style.FontSize * 2).ToString() };
        runProperties.Bold = style.Bold ? new W.Bold() : null;
        runProperties.Italic = style.Italic ? new W.Italic() : null;
        runProperties.Color = new W.Color { Val = style.ColorHex };
    }

    private static void ApplyWordParagraphMarkProperties(W.ParagraphMarkRunProperties markProperties, TextStyle style)
    {
        markProperties.RemoveAllChildren();
        markProperties.Append(new W.RunFonts { Ascii = style.FontName, HighAnsi = style.FontName, ComplexScript = style.FontName });
        markProperties.Append(new W.FontSize { Val = (style.FontSize * 2).ToString() });
        if (style.Bold)
        {
            markProperties.Append(new W.Bold());
        }

        if (style.Italic)
        {
            markProperties.Append(new W.Italic());
        }

        markProperties.Append(new W.Color { Val = style.ColorHex });
    }

    private static void ApplyStyleOptions(W.Style style, JsonObject options, W.StyleValues styleType)
    {
        style.Type = styleType;

        if (!style.Elements<W.UnhideWhenUsed>().Any())
        {
            style.Append(new W.UnhideWhenUsed());
        }

        if (!style.Elements<W.PrimaryStyle>().Any())
        {
            style.Append(new W.PrimaryStyle());
        }

        if (!style.Elements<W.UIPriority>().Any())
        {
            style.Append(new W.UIPriority { Val = 99 });
        }

        var runProperties = style.StyleRunProperties ??= new W.StyleRunProperties();

        var beforePt = options.TryGetPropertyValue("beforePt", out var beforeNode) && beforeNode is not null ? beforeNode.GetValue<int>() : 0;
        var afterPt = options.TryGetPropertyValue("afterPt", out var afterNode) && afterNode is not null ? afterNode.GetValue<int>() : 0;
        var lineSpacing = options.TryGetPropertyValue("lineSpacing", out var lineSpacingNode) && lineSpacingNode is not null ? lineSpacingNode.GetValue<double>() : 0;

        if (styleType == W.StyleValues.Character && (beforePt > 0 || afterPt > 0 || lineSpacing > 0))
        {
            throw new InvalidOperationException("Character styles do not support paragraph spacing options.");
        }

        if (options.TryGetPropertyValue("basedOn", out var basedOnNode) && basedOnNode is not null)
        {
            style.BasedOn = new W.BasedOn { Val = basedOnNode.GetValue<string>() };
        }

        if (options.TryGetPropertyValue("nextStyle", out var nextStyleNode) && nextStyleNode is not null)
        {
            if (styleType != W.StyleValues.Paragraph)
            {
                throw new InvalidOperationException("Only paragraph styles support nextStyle.");
            }

            style.NextParagraphStyle = new W.NextParagraphStyle { Val = nextStyleNode.GetValue<string>() };
        }

        if (options.TryGetPropertyValue("fontName", out var fontNameNode) && fontNameNode is not null)
        {
            var fontName = fontNameNode.GetValue<string>();
            runProperties.RunFonts = new W.RunFonts { Ascii = fontName, HighAnsi = fontName, ComplexScript = fontName };
        }

        if (options.TryGetPropertyValue("fontSize", out var fontSizeNode) && fontSizeNode is not null)
        {
            runProperties.FontSize = new W.FontSize { Val = (fontSizeNode.GetValue<int>() * 2).ToString() };
        }

        if (options.TryGetPropertyValue("bold", out var boldNode) && boldNode is not null)
        {
            runProperties.Bold = boldNode.GetValue<bool>() ? new W.Bold() : null;
        }

        if (options.TryGetPropertyValue("italic", out var italicNode) && italicNode is not null)
        {
            runProperties.Italic = italicNode.GetValue<bool>() ? new W.Italic() : null;
        }

        if (options.TryGetPropertyValue("colorHex", out var colorNode) && colorNode is not null)
        {
            var color = colorNode.GetValue<string>();
            if (!IsValidHexColor(color))
            {
                throw new InvalidOperationException("Style colorHex must be a 6-digit hex value.");
            }

            runProperties.Color = new W.Color { Val = color };
        }

        if (beforePt > 0 || afterPt > 0 || lineSpacing > 0)
        {
            var paragraphProperties = style.StyleParagraphProperties ??= new W.StyleParagraphProperties();
            paragraphProperties.SpacingBetweenLines ??= new W.SpacingBetweenLines();
            if (beforePt >= 0)
            {
                paragraphProperties.SpacingBetweenLines.Before = (beforePt * 20).ToString();
            }

            if (afterPt >= 0)
            {
                paragraphProperties.SpacingBetweenLines.After = (afterPt * 20).ToString();
            }

            if (lineSpacing > 0)
            {
                paragraphProperties.SpacingBetweenLines.Line = ((int)Math.Round(lineSpacing * 240)).ToString();
                paragraphProperties.SpacingBetweenLines.LineRule = W.LineSpacingRuleValues.Auto;
            }
        }
    }

    private static ListNumberingDefinitions EnsureStructuredListNumbering(WordprocessingDocument document, IReadOnlyCollection<WordListItem> items)
    {
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
        var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new W.Numbering();
        var numbering = numberingPart.Numbering;

        const int numberedAbstractId = 901;
        const int bulletedAbstractId = 902;
        var numberedNumberingId = GetNextWordNumberingInstanceId(numbering);
        var bulletedNumberingId = numberedNumberingId + 1;

        var maxLevel = items.Max(x => x.Level);
        EnsureAbstractNumbering(numbering, numberedAbstractId, maxLevel, isNumbered: true, items);
        EnsureAbstractNumbering(numbering, bulletedAbstractId, maxLevel, isNumbered: false, items);

        EnsureNumberingInstance(numbering, numberedNumberingId, numberedAbstractId);
        EnsureNumberingInstance(numbering, bulletedNumberingId, bulletedAbstractId);

        numbering.Save();
        return new ListNumberingDefinitions(numberedNumberingId, bulletedNumberingId);
    }

    private static void EnsureAbstractNumbering(W.Numbering numbering, int abstractId, int maxLevel, bool isNumbered, IReadOnlyCollection<WordListItem> items)
    {
        var existing = numbering.Elements<W.AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == abstractId);
        existing?.Remove();

        var abstractNum = new W.AbstractNum { AbstractNumberId = abstractId };
        for (var level = 0; level <= maxLevel; level++)
        {
            var style = isNumbered
                ? ResolveNumberStyle(items, level)
                : ResolveBulletStyle(items, level);
            abstractNum.Append(BuildListLevel(level, style, isNumbered));
        }

        numbering.Append(abstractNum);
    }

    private static void EnsureNumberingInstance(W.Numbering numbering, int numberingId, int abstractId)
    {
        var existing = numbering.Elements<W.NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numberingId);
        existing?.Remove();
        numbering.Append(new W.NumberingInstance(new W.AbstractNumId { Val = abstractId }) { NumberID = numberingId });
    }

    private static int GetNextWordNumberingInstanceId(W.Numbering numbering)
    {
        var maxExisting = numbering.Elements<W.NumberingInstance>()
            .Select(n => n.NumberID?.Value)
            .Where(v => v.HasValue)
            .Select(v => v!.Value)
            .DefaultIfEmpty(0)
            .Max();

        return Math.Max(1000, maxExisting + 1);
    }

    private static W.Level BuildListLevel(int level, string style, bool isNumbered)
    {
        var left = 720 * (level + 1);
        var hanging = 360;
        var numberingFormat = isNumbered ? ResolveNumberFormat(style) : W.NumberFormatValues.Bullet;
        var levelText = isNumbered ? ResolveNumberLevelText(style, level) : ResolveBulletGlyph(style);

        return new W.Level(
            new W.StartNumberingValue { Val = 1 },
            new W.NumberingFormat { Val = numberingFormat },
            new W.LevelText { Val = levelText },
            new W.LevelJustification { Val = W.LevelJustificationValues.Left },
            new W.ParagraphProperties(new W.Indentation { Left = left.ToString(), Hanging = hanging.ToString() }),
            new W.RunProperties(new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }))
        { LevelIndex = level };
    }

    private static string ResolveNumberStyle(IReadOnlyCollection<WordListItem> items, int level)
    {
        return items.FirstOrDefault(x => x.Level == level && NormalizeListKind(x.Kind) == "numbered")?.NumberStyle
            ?.Trim().ToLowerInvariant()
            ?? "decimal-dot";
    }

    private static string ResolveBulletStyle(IReadOnlyCollection<WordListItem> items, int level)
    {
        return items.FirstOrDefault(x => x.Level == level && NormalizeListKind(x.Kind) == "bulleted")?.BulletStyle
            ?.Trim().ToLowerInvariant()
            ?? "disc";
    }

    private static string NormalizeListKind(string kind)
    {
        var normalized = kind?.Trim().ToLowerInvariant();
        return normalized is "numbered" or "ordered" ? "numbered" : "bulleted";
    }

    private static W.NumberFormatValues ResolveNumberFormat(string style)
    {
        return style switch
        {
            "upper-alpha-dot" => W.NumberFormatValues.UpperLetter,
            "lower-alpha-paren" => W.NumberFormatValues.LowerLetter,
            "upper-roman-dot" => W.NumberFormatValues.UpperRoman,
            "lower-roman-dot" => W.NumberFormatValues.LowerRoman,
            _ => W.NumberFormatValues.Decimal
        };
    }

    private static string ResolveNumberLevelText(string style, int level)
    {
        var marker = $"%{level + 1}";
        return style switch
        {
            "lower-alpha-paren" => $"{marker})",
            "decimal-paren" => $"{marker})",
            _ => $"{marker}."
        };
    }

    private static string ResolveBulletGlyph(string style)
    {
        return style switch
        {
            "circle" => "o",
            "square" => "■",
            "diamond" => "◆",
            "dash" => "-",
            "arrow" => "➤",
            "check" => "✓",
            _ => "•"
        };
    }

    private static List<WordParagraphReference> GetWordParagraphReferences(W.Body body)
    {
        var tables = body.Elements<W.Table>().ToList();
        var references = new List<WordParagraphReference>();
        var allIndex = 0;
        var bodyIndex = 0;

        foreach (var paragraph in body.Descendants<W.Paragraph>())
        {
            allIndex++;
            int? bodyParagraphIndex = null;
            if (ReferenceEquals(paragraph.Parent, body))
            {
                bodyIndex++;
                bodyParagraphIndex = bodyIndex;
            }

            var table = paragraph.Ancestors<W.Table>().FirstOrDefault();
            int? tableIndex = null;
            int? rowIndex = null;
            int? columnIndex = null;

            if (table is not null)
            {
                tableIndex = tables.FindIndex(t => ReferenceEquals(t, table)) + 1;
                var row = paragraph.Ancestors<W.TableRow>().FirstOrDefault();
                var cell = paragraph.Ancestors<W.TableCell>().FirstOrDefault();
                if (row is not null)
                {
                    rowIndex = table.Elements<W.TableRow>().TakeWhile(r => !ReferenceEquals(r, row)).Count() + 1;
                }

                if (cell is not null && row is not null)
                {
                    columnIndex = row.Elements<W.TableCell>().TakeWhile(c => !ReferenceEquals(c, cell)).Count() + 1;
                }
            }

            references.Add(new WordParagraphReference(
                paragraph,
                paragraph.InnerText,
                allIndex,
                bodyParagraphIndex,
                table is not null,
                tableIndex,
                rowIndex,
                columnIndex));
        }

        return references;
    }

    private static List<object> BuildWordTableSummaries(W.Body body)
    {
        var tables = body.Elements<W.Table>().ToList();
        var summaries = new List<object>(tables.Count);
        for (var i = 0; i < tables.Count; i++)
        {
            var table = tables[i];
            var rows = table.Elements<W.TableRow>().ToList();
            var columns = rows.FirstOrDefault()?.Elements<W.TableCell>().Count() ?? 0;
            var cells = rows
                .SelectMany((row, rowIndex) => row.Elements<W.TableCell>().Select((cell, colIndex) => new
                {
                    rowIndex = rowIndex + 1,
                    columnIndex = colIndex + 1,
                    text = string.Join("\n", cell.Elements<W.Paragraph>().Select(p => p.InnerText)),
                    preview = TrimPreview(string.Join(" ", cell.Elements<W.Paragraph>().Select(p => p.InnerText)))
                }))
                .ToArray();

            summaries.Add(new
            {
                tableIndex = i + 1,
                rowCount = rows.Count,
                columnCount = columns,
                cells
            });
        }

        return summaries;
    }

    private static W.Table GetWordTableByIndex(W.Body body, int tableIndex)
    {
        var tables = body.Elements<W.Table>().ToList();
        if (tableIndex > tables.Count)
        {
            throw new InvalidOperationException($"Table index {tableIndex} is out of range. Valid range is 1..{tables.Count}.");
        }

        return tables[tableIndex - 1];
    }

    private static string BuildWordBodyParagraphRangeMessage(W.Body body, int requestedIndex, int bodyParagraphCount)
    {
        var allParagraphCount = body.Descendants<W.Paragraph>().Count();
        var tableParagraphs = allParagraphCount - bodyParagraphCount;
        return $"Paragraph index {requestedIndex} is out of range. Valid range is 1..{bodyParagraphCount} for body-level paragraphs. This document also contains {tableParagraphs} table-cell paragraphs that are not addressable by paragraph index tools.";
    }

    private static int? GetHeadingLevel(W.Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrWhiteSpace(styleId))
        {
            return null;
        }

        var normalized = styleId.Trim();
        if (normalized.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(normalized[7..], out var level)
            && level >= 1
            && level <= 6)
        {
            return level;
        }

        return null;
    }

    private static Dictionary<int, string> BuildListKindMap(MainDocumentPart mainPart)
    {
        var map = new Dictionary<int, string>();
        var numbering = mainPart.NumberingDefinitionsPart?.Numbering;
        if (numbering is null)
        {
            return map;
        }

        var abstractById = numbering.Elements<W.AbstractNum>()
            .Where(x => x.AbstractNumberId?.Value is not null)
            .ToDictionary(x => x.AbstractNumberId!.Value!, x => x);

        foreach (var instance in numbering.Elements<W.NumberingInstance>())
        {
            var numId = instance.NumberID?.Value;
            var abstractId = instance.AbstractNumId?.Val?.Value;
            if (!numId.HasValue || !abstractId.HasValue || !abstractById.TryGetValue(abstractId.Value, out var abstractNum))
            {
                continue;
            }

            var firstLevel = abstractNum.Elements<W.Level>().OrderBy(l => l.LevelIndex?.Value ?? 0).FirstOrDefault();
            var format = firstLevel?.NumberingFormat?.Val?.Value;
            var resolvedNumId = numId.Value;
            map[resolvedNumId] = format == W.NumberFormatValues.Bullet ? "bulleted" : "numbered";
        }

        return map;
    }

    private static string ResolveListKind(IReadOnlyDictionary<int, string> listKindByNumberingId, int numberingId)
    {
        return listKindByNumberingId.TryGetValue(numberingId, out var kind) ? kind : "numbered";
    }

    private static IEnumerable<WordRunSegment> BuildParagraphRunSegments(W.Paragraph paragraph)
    {
        var offset = 0;
        foreach (var run in paragraph.Descendants<W.Run>())
        {
            var text = run.InnerText;
            if (string.IsNullOrEmpty(text))
            {
                continue;
            }

            yield return new WordRunSegment(run, offset, text.Length, text);
            offset += text.Length;
        }
    }

    private static List<(int Start, int Length)> FindWholeWordMatches(string text, string query, StringComparison comparison)
    {
        var matches = new List<(int Start, int Length)>();
        var startIndex = 0;
        while (true)
        {
            var index = text.IndexOf(query, startIndex, comparison);
            if (index < 0)
            {
                break;
            }

            var endIndex = index + query.Length;
            var leftBoundary = index == 0 || !IsWordChar(text[index - 1]);
            var rightBoundary = endIndex == text.Length || !IsWordChar(text[endIndex]);
            if (leftBoundary && rightBoundary)
            {
                matches.Add((index, query.Length));
            }

            startIndex = index + Math.Max(1, query.Length);
        }

        return matches;
    }

    private static bool IsWordChar(char c) => char.IsLetterOrDigit(c) || c == '_';

    private static List<(int Start, int Length)> FindAllMatches(string text, string query, StringComparison comparison)
    {
        var matches = new List<(int Start, int Length)>();
        var startIndex = 0;
        while (true)
        {
            var index = text.IndexOf(query, startIndex, comparison);
            if (index < 0) break;
            matches.Add((index, query.Length));
            startIndex = index + Math.Max(1, query.Length);
        }
        return matches;
    }

    private static W.Style? ResolveWordStyle(MainDocumentPart mainPart, string styleName)
    {
        var styles = mainPart.StyleDefinitionsPart?.Styles?.Elements<W.Style>() ?? [];
        return styles.FirstOrDefault(s =>
            string.Equals(s.StyleId?.Value, styleName, StringComparison.OrdinalIgnoreCase)
            || string.Equals(s.StyleName?.Val?.Value, styleName, StringComparison.OrdinalIgnoreCase));
    }

    private static Dictionary<string, string> BuildWordStyleNameMap(MainDocumentPart? mainPart)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var definition in WordBuiltInStyles)
        {
            map[definition.StyleId] = definition.Name;
        }

        var styles = mainPart?.StyleDefinitionsPart?.Styles?.Elements<W.Style>() ?? [];
        foreach (var style in styles)
        {
            var styleId = style.StyleId?.Value;
            if (string.IsNullOrWhiteSpace(styleId))
            {
                continue;
            }

            map[styleId] = style.StyleName?.Val?.Value ?? styleId;
        }

        return map;
    }

    private static void SplitRunAndApplyCharacterStyle(W.Run run, string text, int matchStart, int matchLength, string styleId)
    {
        var beforeText = text[..matchStart];
        var matchText = text.Substring(matchStart, matchLength);
        var afterText = text[(matchStart + matchLength)..];

        if (beforeText.Length > 0)
        {
            run.InsertBeforeSelf(BuildRunCloneWithText(run, beforeText));
        }

        run.InsertBeforeSelf(BuildRunCloneWithText(run, matchText, styleId));

        if (afterText.Length > 0)
        {
            run.InsertBeforeSelf(BuildRunCloneWithText(run, afterText));
        }

        run.Remove();
    }

    private static W.Run BuildRunCloneWithText(W.Run template, string text, string? runStyleId = null)
    {
        var clone = new W.Run();
        if (template.RunProperties is not null)
        {
            clone.RunProperties = (W.RunProperties)template.RunProperties.CloneNode(true);
        }

        if (!string.IsNullOrWhiteSpace(runStyleId))
        {
            clone.RunProperties ??= new W.RunProperties();
            clone.RunProperties.RunStyle = new W.RunStyle { Val = runStyleId };
        }

        var textNode = new W.Text(text);
        if (text.StartsWith(' ') || text.EndsWith(' '))
        {
            textNode.Space = SpaceProcessingModeValues.Preserve;
        }

        clone.Append(textNode);
        return clone;
    }

    private static string ResolveWordStyleType(W.Style style) => ResolveWordStyleType(style.Type?.Value);

    private static string ResolveWordStyleType(W.StyleValues? type)
    {
        if (type == W.StyleValues.Paragraph) return "paragraph";
        if (type == W.StyleValues.Character) return "character";
        if (type == W.StyleValues.Table) return "table";
        if (type == W.StyleValues.Numbering) return "numbering";
        return string.Empty;
    }

    private static W.StyleValues ParseWordStyleType(string type)
    {
        return type.Trim().ToLowerInvariant() switch
        {
            "paragraph" => W.StyleValues.Paragraph,
            "character" => W.StyleValues.Character,
            _ => throw new InvalidOperationException("Style type must be one of: paragraph, character.")
        };
    }

    private static string SanitizeStyleId(string styleName)
    {
        var cleaned = new string(styleName.Where(char.IsLetterOrDigit).ToArray());
        return string.IsNullOrWhiteSpace(cleaned) ? $"Custom{Guid.NewGuid():N}" : cleaned;
    }

    private static string TrimPreview(string text, int maxLength = 120)
    {
        var normalized = string.Join(" ", (text ?? string.Empty).Split('\n', '\r').Select(x => x.Trim()).Where(x => x.Length > 0));
        if (normalized.Length <= maxLength)
        {
            return normalized;
        }

        return normalized[..maxLength] + "...";
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

    private static bool IsValidHexColor(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length != 6)
        {
            return false;
        }

        return value.All(Uri.IsHexDigit);
    }

    public static void ValidateTextStyle(TextStyle style)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(style.FontName);
        if (style.FontSize < 8 || style.FontSize > 96)
        {
            throw new InvalidOperationException("Font size must be between 8 and 96.");
        }

        if (!IsValidHexColor(style.ColorHex))
        {
            throw new InvalidOperationException("Color must be a 6-digit hex value like FF0000.");
        }
    }

    private static IReadOnlyList<WordBuiltInStyleDefinition> BuildWordBuiltInStyles()
    {
        var styles = new List<WordBuiltInStyleDefinition>
        {
            new("Normal", "Normal", "paragraph", IsDefault: true, UiPriority: 0, IsPrimary: true),
            new("NoSpacing", "No Spacing", "paragraph", BasedOn: "Normal", NextStyle: "Normal", UiPriority: 1, IsPrimary: true)
        };

        for (var level = 1; level <= 9; level++)
        {
            styles.Add(new WordBuiltInStyleDefinition($"Heading{level}", $"Heading {level}", "paragraph", BasedOn: "Normal", NextStyle: "Normal", UiPriority: 9, IsPrimary: true));
        }

        styles.AddRange(
        [
            new("Title", "Title", "paragraph", BasedOn: "Normal", NextStyle: "Normal", UiPriority: 10, IsPrimary: true),
            new("Subtitle", "Subtitle", "paragraph", BasedOn: "Normal", NextStyle: "Normal", UiPriority: 11, IsPrimary: true),
            new("SubtleEmphasis", "Subtle Emphasis", "character", UiPriority: 19, IsPrimary: true),
            new("Emphasis", "Emphasis", "character", UiPriority: 20, IsPrimary: true),
            new("IntenseEmphasis", "Intense Emphasis", "character", UiPriority: 21, IsPrimary: true),
            new("Strong", "Strong", "character", UiPriority: 22, IsPrimary: true),
            new("Quote", "Quote", "paragraph", BasedOn: "Normal", NextStyle: "Normal", UiPriority: 29, IsPrimary: true),
            new("IntenseQuote", "Intense Quote", "paragraph", BasedOn: "Quote", NextStyle: "Normal", UiPriority: 30, IsPrimary: true),
            new("SubtleReference", "Subtle Reference", "character", UiPriority: 31, IsPrimary: true),
            new("IntenseReference", "Intense Reference", "character", UiPriority: 32, IsPrimary: true),
            new("BookTitle", "Book Title", "character", UiPriority: 33, IsPrimary: true),
            new("ListParagraph", "List Paragraph", "paragraph", BasedOn: "Normal", NextStyle: "Normal", UiPriority: 34, IsPrimary: true),
            new("Caption", "Caption", "paragraph", BasedOn: "Normal", NextStyle: "Normal", UiPriority: 35, IsPrimary: true),
            new("Header", "Header", "paragraph", BasedOn: "Normal", NextStyle: "Normal", IsPrimary: false),
            new("Footer", "Footer", "paragraph", BasedOn: "Normal", NextStyle: "Normal", IsPrimary: false),
            new("FootnoteText", "Footnote Text", "paragraph", BasedOn: "Normal", NextStyle: "FootnoteText", IsPrimary: false),
            new("EndnoteText", "Endnote Text", "paragraph", BasedOn: "Normal", NextStyle: "EndnoteText", IsPrimary: false)
        ]);

        return styles;
    }

    // -------------------------------------------------------------------------
    // Private record types
    // -------------------------------------------------------------------------

    private sealed record WordBuiltInStyleDefinition(
        string StyleId,
        string Name,
        string Type,
        bool IsDefault = false,
        string? BasedOn = null,
        string? NextStyle = null,
        int? UiPriority = null,
        bool IsPrimary = false,
        bool UnhideWhenUsed = true);

    private sealed record WordRunSegment(W.Run Run, int Start, int Length, string Text);

    private sealed record ListNumberingDefinitions(int NumberedNumberingId, int BulletedNumberingId);

    private sealed record WordParagraphReference(
        W.Paragraph Paragraph,
        string Text,
        int AllParagraphIndex,
        int? BodyParagraphIndex,
        bool InTable,
        int? TableIndex,
        int? RowIndex,
        int? ColumnIndex);
}
