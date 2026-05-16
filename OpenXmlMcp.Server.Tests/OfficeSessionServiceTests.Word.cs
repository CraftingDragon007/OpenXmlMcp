using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlMcp.Server.Services;

namespace OpenXmlMcp.Server.Tests;

public partial class OfficeSessionServiceTests
{
    [Fact]
    public void WordSession_CreateAppendFind_Works()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Hello MCP Word");

            var structure = service.ListStructure(sessionId);
            Assert.Contains("paragraphCount", structure);

            var findResult = service.FindText(sessionId, "mcp word");
            Assert.Contains("matchCount", findResult);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordSession_AddTable_IncreasesTableCount()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 2, 3);

            var structure = service.ListStructure(sessionId);
            Assert.Contains("\"tableCount\":1", structure, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordTableCell_SetAndGet_Works()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 2, 2);
            service.WordSetTableCell(sessionId, 1, 2, 1, "R2C1");

            var value = service.WordGetTableCell(sessionId, 1, 2, 1);
            Assert.Equal("R2C1", value);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordStructure_ContainsElementsAndTableSummaries()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddHeading(sessionId, 1, "Title");
            service.WordAppendParagraph(sessionId, "Intro");
            service.WordAddTable(sessionId, 1, 1);
            service.WordSetTableCell(sessionId, 1, 1, 1, "Cell text");

            var structure = service.ListStructure(sessionId);
            Assert.Contains("\"elements\"", structure, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"tables\"", structure, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"bodyParagraphCount\"", structure, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Cell text", structure, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordFindText_InTable_ReportsAddressabilityMetadata()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Body anchor");
            service.WordAddTable(sessionId, 1, 1);
            service.WordSetTableCell(sessionId, 1, 1, 1, "Table anchor");

            var result = service.FindText(sessionId, "anchor");
            Assert.Contains("\"addressableByParagraphTools\":true", result, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"addressableByParagraphTools\":false", result, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"tableIndex\":1", result, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointBulletSlide_UsesRealBulletParagraphs()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddBulletSlide(sessionId, "Sprint", "Task A\nTask B");
            service.CloseDocument(sessionId);

            using var presentation = PresentationDocument.Open(filePath, false);
            var slidePart = presentation.PresentationPart!.SlideParts.First();
            Assert.NotNull(slidePart.Slide);
            var bulletParagraphs = slidePart.Slide!
                .Descendants<DocumentFormat.OpenXml.Drawing.ParagraphProperties>()
                .Count(p => p.Descendants<DocumentFormat.OpenXml.Drawing.CharacterBullet>().Any(b => b.Char?.Value == "•"));

            Assert.True(bulletParagraphs >= 2);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointAddSlide_NumberedBody_UsesAutoNumberedParagraphs()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Plan", "Kickoff\nScope\nDelivery", "numbered");
            service.CloseDocument(sessionId);

            using var presentation = PresentationDocument.Open(filePath, false);
            var slidePart = presentation.PresentationPart!.SlideParts.First();
            Assert.NotNull(slidePart.Slide);
            var numberedParagraphs = slidePart.Slide!
                .Descendants<DocumentFormat.OpenXml.Drawing.ParagraphProperties>()
                .Count(p => p.Descendants<DocumentFormat.OpenXml.Drawing.AutoNumberedBullet>().Any());

            Assert.True(numberedParagraphs >= 3);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordInsertParagraphAt_UsesOneBasedIndex()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "first");
            service.WordAppendParagraph(sessionId, "third");
            service.WordInsertParagraphAt(sessionId, 2, "second");

            var matches = service.FindText(sessionId, "second");
            Assert.Contains("\"index\":2", matches, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordReplaceText_ReturnsReplacementCount()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "alpha beta alpha");

            var payload = service.WordReplaceText(sessionId, "alpha", "gamma", matchCase: false);

            Assert.Contains("\"replacementCount\":2", payload, StringComparison.OrdinalIgnoreCase);
            var result = service.FindText(sessionId, "gamma");
            Assert.Contains("\"matchCount\":1", result, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordAddHeading_AndBulletedList_AreSearchable()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddHeading(sessionId, 2, "Roadmap");
            service.WordAddBulletedList(sessionId, "Item A\nItem B");

            var headingResult = service.FindText(sessionId, "Roadmap");
            var bulletResult = service.FindText(sessionId, "Item B");
            Assert.Contains("\"matchCount\":1", headingResult, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"matchCount\":1", bulletResult, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointSetTextStyle_BodySlotStylesAllBodyParagraphs()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Plan", "Kickoff\nScope\nDelivery", "bulleted");
            service.PowerPointSetTextStyle(sessionId, 1, 1, "Aptos", 22, bold: false, italic: false, colorHex: "374151");
            service.CloseDocument(sessionId);

            using var presentation = PresentationDocument.Open(filePath, false);
            var slidePart = presentation.PresentationPart!.SlideParts.First();
            Assert.NotNull(slidePart.Slide);

            var bodySizes = slidePart.Slide!
                .Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>()
                .Skip(1)
                .Select(p => p.Descendants<DocumentFormat.OpenXml.Drawing.RunProperties>().FirstOrDefault()?.FontSize?.Value)
                .ToList();

            Assert.True(bodySizes.Count >= 3);
            Assert.All(bodySizes, size => Assert.Equal(2200, size));
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordSetParagraphStyle_AppliesFormatting()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Styled text");
            service.WordSetParagraphStyle(sessionId, 1, "Calibri", 20, bold: true, italic: true, colorHex: "FF0000");
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var runProperties = document.MainDocumentPart?
                .Document?
                .Body?
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .First()
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Run>()
                .First()
                .RunProperties;

            Assert.NotNull(runProperties);
            Assert.NotNull(runProperties!.Bold);
            Assert.NotNull(runProperties.Italic);
            Assert.Equal("FF0000", runProperties.Color?.Val?.Value);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordAddNumberedList_UsesDecimalDotByDefault()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddNumberedList(sessionId, "One\nTwo");
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var paragraph = document.MainDocumentPart?.Document?.Body?.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First();
            var numberingProps = paragraph?.ParagraphProperties?.NumberingProperties;

            Assert.NotNull(numberingProps);
            Assert.Equal(0, numberingProps!.NumberingLevelReference?.Val?.Value);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAddNumberedList_SeparateCallsUseDifferentNumberingInstances()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddNumberedList(sessionId, "One\nTwo");
            service.WordAddHeading(sessionId, 2, "Break");
            service.WordAddNumberedList(sessionId, "Three\nFour");
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var numberedParagraphs = document.MainDocumentPart?.Document?.Body?
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .Where(p => p.ParagraphProperties?.NumberingProperties?.NumberingId is not null)
                .ToList() ?? [];

            Assert.True(numberedParagraphs.Count >= 4);

            var firstListNumId = numberedParagraphs[0].ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value;
            var secondListNumId = numberedParagraphs[2].ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value;
            Assert.NotEqual(firstListNumId, secondListNumId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordListStructure_DistinguishesNumberedAndBulletedKinds()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddNumberedList(sessionId, "One\nTwo");
            service.WordAddBulletedList(sessionId, "Alpha\nBeta");

            var structureJson = service.ListStructure(sessionId);
            using var doc = JsonDocument.Parse(structureJson);
            var elements = doc.RootElement.GetProperty("elements").EnumerateArray().ToArray();

            var numbered = elements.Where(e => e.TryGetProperty("text", out var t) && (t.GetString() == "One" || t.GetString() == "Two")).ToArray();
            var bulleted = elements.Where(e => e.TryGetProperty("text", out var t) && (t.GetString() == "Alpha" || t.GetString() == "Beta")).ToArray();

            Assert.All(numbered, e => Assert.Equal("numbered", e.GetProperty("listKind").GetString()));
            Assert.All(bulleted, e => Assert.Equal("bulleted", e.GetProperty("listKind").GetString()));

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordAddStructuredList_SupportsNestedMixedKinds()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var itemsJson = "[" +
                "{\"text\":\"Top 1\",\"level\":0,\"kind\":\"numbered\"}," +
                "{\"text\":\"Nested A\",\"level\":1,\"kind\":\"bulleted\",\"bulletStyle\":\"disc\"}," +
                "{\"text\":\"Nested B\",\"level\":1,\"kind\":\"bulleted\",\"bulletStyle\":\"square\"}" +
                "]";
            service.WordAddStructuredList(sessionId, itemsJson);
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var paragraphs = document.MainDocumentPart?.Document?.Body?.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList()
                ?? throw new InvalidOperationException("Paragraphs missing.");

            Assert.Equal(3, paragraphs.Count);
            Assert.Equal(0, paragraphs[0].ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value);
            Assert.Equal(1, paragraphs[1].ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value);
            Assert.Equal(1, paragraphs[2].ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordAppendParagraph_AppliesDefaultSpacing()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Spacing test");
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var spacing = document.MainDocumentPart?
                .Document?
                .Body?
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .First()
                .ParagraphProperties?
                .SpacingBetweenLines;

            Assert.NotNull(spacing);
            Assert.Equal("160", spacing!.After?.Value);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordSetDocumentSpacingPreset_Comfortable_UpdatesParagraphSpacing()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "A");
            service.WordAppendParagraph(sessionId, "B");
            service.WordSetDocumentSpacingPreset(sessionId, "comfortable");
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var paragraphs = document.MainDocumentPart?
                .Document?
                .Body?
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .ToList() ?? [];

            Assert.NotEmpty(paragraphs);
            Assert.Equal("200", paragraphs[0].ParagraphProperties?.SpacingBetweenLines?.After?.Value);
            Assert.Equal("312", paragraphs[0].ParagraphProperties?.SpacingBetweenLines?.Line?.Value);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordInsertParagraphAfterText_InsertsAtExpectedPosition()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Anchor paragraph");
            var payload = service.WordInsertParagraphAfterText(sessionId, "Anchor", "Inserted paragraph");

            Assert.Contains("\"insertedIndex\":2", payload, StringComparison.OrdinalIgnoreCase);
            var find = service.FindText(sessionId, "Inserted paragraph");
            Assert.Contains("\"index\":2", find, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordInsertTextAfterText_InsertsInline()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Hello world");
            var payload = service.WordInsertTextAfterText(sessionId, "Hello", " dear");

            Assert.Contains("\"changed\":true", payload, StringComparison.OrdinalIgnoreCase);
            var find = service.FindText(sessionId, "Hello dear world");
            Assert.Contains("\"matchCount\":1", find, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordCreateOrUpdateStyle_AndApplyStyleByName_Work()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Styled paragraph");
            service.WordCreateOrUpdateStyle(sessionId, "MyBlueStyle", "{\"fontName\":\"Calibri\",\"fontSize\":14,\"colorHex\":\"1F4E79\",\"afterPt\":8,\"lineSpacing\":1.15}");
            service.WordApplyStyleByName(sessionId, 1, "MyBlueStyle");

            var styles = service.WordListStyles(sessionId);
            Assert.Contains("MyBlueStyle", styles, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordListStyles_IncludesBuiltInDefaultsOnNewDocument()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var stylesPayload = service.WordListStyles(sessionId);

            Assert.Contains("\"styleId\":\"Normal\"", stylesPayload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"styleId\":\"Heading1\"", stylesPayload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"styleId\":\"NoSpacing\"", stylesPayload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"styleId\":\"Strong\"", stylesPayload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordApplyStyleByName_ResolvesEnglishBuiltInName()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Heading paragraph");
            service.WordApplyStyleByName(sessionId, 1, "Heading 2");

            var infoPayload = service.WordGetParagraphInfo(sessionId, 1);
            Assert.Contains("\"styleId\":\"Heading2\"", infoPayload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordCreateOrUpdateStyle_CreatesCharacterStyle()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var payload = service.WordCreateOrUpdateStyle(sessionId, "CalloutInline", "{\"type\":\"character\",\"fontName\":\"Calibri\",\"bold\":true}");

            Assert.Contains("\"type\":\"character\"", payload, StringComparison.OrdinalIgnoreCase);

            var stylesPayload = service.WordListStyles(sessionId);
            Assert.Contains("\"styleId\":\"CalloutInline\"", stylesPayload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"type\":\"character\"", stylesPayload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordApplyCharacterStyleToText_AndListParagraphRuns_Work()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "alpha beta gamma");
            service.WordApplyCharacterStyleToText(sessionId, "beta", "Strong");

            var runPayload = service.WordListParagraphRuns(sessionId, 1);
            Assert.Contains("\"runCount\":3", runPayload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"styleId\":\"Strong\"", runPayload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordApplyCharacterStyleToText_UsesWholeWordMatching()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Parameters Parameter");
            service.WordApplyCharacterStyleToText(sessionId, "Parameter", "Strong");

            var runPayload = service.WordListParagraphRuns(sessionId, 1);
            using var parsed = JsonDocument.Parse(runPayload);
            var runs = parsed.RootElement.GetProperty("runs").EnumerateArray().ToArray();
            Assert.Equal("Parameters ", runs[0].GetProperty("text").GetString());
            Assert.Equal(string.Empty, runs[0].GetProperty("styleId").GetString());
            Assert.Equal("Parameter", runs[1].GetProperty("text").GetString());
            Assert.Equal("Strong", runs[1].GetProperty("styleId").GetString());

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void WordGetParagraphInfo_ReturnsStyleAndSpacingFields()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Info paragraph");
            service.WordSetParagraphSpacing(sessionId, 1, 4, 10, 1.2);

            var payload = service.WordGetParagraphInfo(sessionId, 1);
            Assert.Contains("\"paragraphIndex\":1", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"spacingBeforeTwips\"", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordApplyCharacterStyleToAll_StylesAllWholeWordMatches()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Parameters Parameter path");
            service.WordAppendParagraph(sessionId, "Another Parameter here");

            var result = service.WordApplyCharacterStyleToAll(
                sessionId,
                "[\"Parameter\",\"path\"]",
                "Strong",
                matchCase: false,
                wholeWord: true);

            using var doc = JsonDocument.Parse(result);
            Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());
            // "Parameter" appears twice (not "Parameters"), "path" appears once = 3 total
            Assert.Equal(3, doc.RootElement.GetProperty("target").GetProperty("totalStyled").GetInt32());

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordApplyCharacterStyleByPattern_StylesPatternMatches()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Use -server and -port flags");
            service.WordAppendParagraph(sessionId, "Set CAQTDM_WEB_PATH and HOME_DIR env vars");

            var r1 = service.WordApplyCharacterStyleByPattern(sessionId, @"-[A-Za-z0-9_]+", "Strong", matchCase: true);
            var r2 = service.WordApplyCharacterStyleByPattern(sessionId, @"[A-Z][A-Z0-9_]{2,}", "IntenseReference", matchCase: true);

            using var d1 = JsonDocument.Parse(r1);
            using var d2 = JsonDocument.Parse(r2);
            Assert.True(d1.RootElement.GetProperty("ok").GetBoolean());
            Assert.Equal(2, d1.RootElement.GetProperty("target").GetProperty("totalStyled").GetInt32());
            Assert.True(d2.RootElement.GetProperty("ok").GetBoolean());

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordInsertAfterHeading_InsertsCorrectly()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddHeading(sessionId, 1, "Introduction");
            service.WordAppendParagraph(sessionId, "Existing body");
            service.WordAddHeading(sessionId, 2, "Details");

            var result = service.WordInsertAfterHeading(sessionId, "Introduction", "Inserted paragraph");

            using var doc = JsonDocument.Parse(result);
            Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());
            Assert.Equal(2, doc.RootElement.GetProperty("target").GetProperty("insertedIndex").GetInt32());

            // Verify position: para after heading is the inserted one
            var info = service.WordGetParagraphInfo(sessionId, 2);
            Assert.Contains("Inserted paragraph", info, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordReplaceSection_ReplacesBodyUnderHeading()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddHeading(sessionId, 1, "Overview");
            service.WordAppendParagraph(sessionId, "Old paragraph one");
            service.WordAppendParagraph(sessionId, "Old paragraph two");
            service.WordAddHeading(sessionId, 1, "Details");

            var result = service.WordReplaceSection(
                sessionId, "Overview",
                "[\"New content paragraph\"]");

            using var doc = JsonDocument.Parse(result);
            Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());
            Assert.Equal(2, doc.RootElement.GetProperty("target").GetProperty("removedCount").GetInt32());
            Assert.Equal(1, doc.RootElement.GetProperty("target").GetProperty("insertedCount").GetInt32());

            // Document should now have: Overview heading, new para, Details heading
            var find = service.FindText(sessionId, "Old paragraph");
            Assert.Contains("\"matchCount\":0", find, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAddTableRow_AppendsAndInsertsRow()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 2, 3);

            // Append a row
            var appendResult = service.WordAddTableRow(sessionId, 1);
            using (var doc1 = JsonDocument.Parse(appendResult))
            {
                Assert.True(doc1.RootElement.GetProperty("ok").GetBoolean());
                Assert.Equal(3, doc1.RootElement.GetProperty("target").GetProperty("rowCount").GetInt32());
                Assert.Equal(3, doc1.RootElement.GetProperty("target").GetProperty("rowIndex").GetInt32());
            }

            // Insert at index 1 (before first row)
            var insertResult = service.WordAddTableRow(sessionId, 1, 1);
            using (var doc2 = JsonDocument.Parse(insertResult))
            {
                Assert.Equal(4, doc2.RootElement.GetProperty("target").GetProperty("rowCount").GetInt32());
                Assert.Equal(1, doc2.RootElement.GetProperty("target").GetProperty("rowIndex").GetInt32());
            }

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordDeleteTableRow_RemovesRow()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 3, 2);
            service.WordSetTableCell(sessionId, 1, 1, 1, "Header");

            var result = service.WordDeleteTableRow(sessionId, 1, 1);
            using var doc = JsonDocument.Parse(result);
            Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());
            Assert.Equal(2, doc.RootElement.GetProperty("target").GetProperty("rowCount").GetInt32());

            // Verify the header cell is gone
            var cell = service.WordGetTableCell(sessionId, 1, 1, 1);
            Assert.NotEqual("Header", cell);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAddTableColumn_AppendsAndInsertsColumn()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 2, 2);

            // Append a column
            var appendResult = service.WordAddTableColumn(sessionId, 1);
            using (var doc1 = JsonDocument.Parse(appendResult))
            {
                Assert.True(doc1.RootElement.GetProperty("ok").GetBoolean());
                Assert.Equal(3, doc1.RootElement.GetProperty("target").GetProperty("columnCount").GetInt32());
                Assert.Equal(3, doc1.RootElement.GetProperty("target").GetProperty("columnIndex").GetInt32());
            }

            // Insert at column 1 (before first column)
            var insertResult = service.WordAddTableColumn(sessionId, 1, 1);
            using (var doc2 = JsonDocument.Parse(insertResult))
            {
                Assert.Equal(4, doc2.RootElement.GetProperty("target").GetProperty("columnCount").GetInt32());
                Assert.Equal(1, doc2.RootElement.GetProperty("target").GetProperty("columnIndex").GetInt32());
            }

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordDeleteTableColumn_RemovesColumnFromAllRows()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 2, 3);
            service.WordSetTableCell(sessionId, 1, 1, 2, "MiddleCol");

            var result = service.WordDeleteTableColumn(sessionId, 1, 2);
            using var doc = JsonDocument.Parse(result);
            Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());
            Assert.Equal(2, doc.RootElement.GetProperty("target").GetProperty("columnCount").GetInt32());

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordMergeTableCells_SetsGridSpanAndRemovesCells()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 2, 4);

            // Merge columns 1-3 in row 1
            var result = service.WordMergeTableCells(sessionId, 1, 1, 1, 3);
            using var doc = JsonDocument.Parse(result);
            Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());
            Assert.Equal(2, doc.RootElement.GetProperty("target").GetProperty("mergedCellCount").GetInt32());

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordDeleteParagraph_RemovesBodyParagraph()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Keep this");
            service.WordAppendParagraph(sessionId, "Delete this");
            service.WordAppendParagraph(sessionId, "Keep this too");

            // Find index of "Delete this" — it was appended after the first default paragraph,
            // so it's paragraph 3 (default + Keep + Delete)
            var findBefore = service.FindText(sessionId, "Delete this");
            Assert.Contains("matchCount\":1", findBefore);

            var result = service.WordDeleteParagraph(sessionId, 2);
            using var doc = JsonDocument.Parse(result);
            Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());
            Assert.Equal(2, doc.RootElement.GetProperty("target").GetProperty("removedIndex").GetInt32());

            var findAfter = service.FindText(sessionId, "Delete this");
            Assert.Contains("matchCount\":0", findAfter);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordDeleteStyle_CustomOnly_ReturnsOrphans()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");

            // Create a custom style and apply it
            service.WordCreateOrUpdateStyle(sessionId, "MyCustomStyle",
                "{\"fontSize\":14,\"colorHex\":\"AA0000\"}");
            service.WordAppendParagraph(sessionId, "Styled paragraph");
            service.WordApplyStyleByName(sessionId, 1, "MyCustomStyle");
            service.WordAppendParagraph(sessionId, "Another paragraph");
            service.WordApplyStyleByName(sessionId, 2, "MyCustomStyle");

            var result = service.WordDeleteStyle(sessionId, "MyCustomStyle");
            using var doc = JsonDocument.Parse(result);
            Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());
            Assert.Equal("MyCustomStyle", doc.RootElement.GetProperty("deletedStyle").GetString());
            Assert.True(doc.RootElement.GetProperty("orphanedReferenceCount").GetInt32() >= 2);

            // Verify the style is gone from listing
            var styles = service.WordListStyles(sessionId);
            Assert.DoesNotContain("MyCustomStyle", styles);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordDeleteStyle_BuiltIn_Throws()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");

            var ex = Assert.Throws<InvalidOperationException>(() =>
                service.WordDeleteStyle(sessionId, "Heading 1"));
            Assert.Contains("built-in", ex.Message, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordUpdateStyle_ModifiesBuiltInStyle()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var result = service.WordUpdateStyle(sessionId, "Heading 1",
                "{\"fontSize\":20,\"colorHex\":\"1F4E79\"}");

            Assert.Contains("\"ok\":true", result, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"styleName\":\"Heading 1\"", result, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAppendMarkdown_ProducesHeadingsListsAndTable()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        var markdown = "## Runtime Flow\n\nThis is a **bold** intro with `code`.\n\n- Alpha\n- Beta\n- Gamma\n\n| Phase | Action |\n|---|---|\n| Start | Init |\n| Stop | Cleanup |";

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var result = service.WordAppendMarkdown(sessionId, markdown);

            Assert.Contains("\"ok\":true", result, StringComparison.OrdinalIgnoreCase);

            var structure = service.ListStructure(sessionId);
            Assert.Contains("heading", structure, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("tableCount", structure, StringComparison.OrdinalIgnoreCase);

            var find = service.FindText(sessionId, "Runtime Flow");
            Assert.Contains("\"matchCount\":1", find, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAppendMarkdown_StrayPipeLinesDoNotCreateTables()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        var markdown = "Install path | optional\nAnother line with | single pipe";

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendMarkdown(sessionId, markdown);

            var structure = service.ListStructure(sessionId);
            Assert.Contains("\"tableCount\":0", structure, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAppendMarkdown_JoinsWrappedLinesIntoSingleParagraph()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        var markdown = "This is a wrapped\nline that should stay\nin one paragraph.\n\nNext paragraph.";

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendMarkdown(sessionId, markdown);
            service.CloseDocument(sessionId);

            using var document = WordprocessingDocument.Open(filePath, false);
            var paragraphs = document.MainDocumentPart?.Document?.Body?
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .Select(p => p.InnerText)
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .ToList() ?? [];

            Assert.Contains("This is a wrapped line that should stay in one paragraph.", paragraphs);
            Assert.Contains("Next paragraph.", paragraphs);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordValidateDocument_DetectsSkippedHeadingLevel()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddHeading(sessionId, 1, "Top");
            service.WordAddHeading(sessionId, 3, "Skipped H2"); // skips level 2

            var result = service.WordValidateDocument(sessionId);

            using var doc = JsonDocument.Parse(result);
            Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());

            var issues = doc.RootElement.GetProperty("issues").EnumerateArray().ToArray();
            Assert.Contains(issues, i => i.GetProperty("code").GetString() == "SkippedHeadingLevel");

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordValidateDocument_DetectsEmptyTableCell()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 2, 2);
            service.WordSetTableCell(sessionId, 1, 1, 1, "Filled");
            // row1col2, row2col1, row2col2 remain empty

            var result = service.WordValidateDocument(sessionId);

            using var doc = JsonDocument.Parse(result);
            var issues = doc.RootElement.GetProperty("issues").EnumerateArray().ToArray();
            Assert.Contains(issues, i => i.GetProperty("code").GetString() == "EmptyTableCell");

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordInsertTableOfContents_InsertsFieldCode()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddHeading(sessionId, 1, "Chapter One");
            service.WordAddHeading(sessionId, 2, "Section 1.1");

            var result = service.WordInsertTableOfContents(sessionId, 1, minLevel: 1, maxLevel: 3);

            Assert.Contains("\"ok\":true", result, StringComparison.OrdinalIgnoreCase);

            // Verify TOC field is present in document XML
            service.CloseDocument(sessionId);

            using var zipFile = System.IO.Compression.ZipFile.OpenRead(filePath);
            var docEntry = zipFile.GetEntry("word/document.xml");
            Assert.NotNull(docEntry);
            using var stream = docEntry!.Open();
            using var reader = new System.IO.StreamReader(stream);
            var xml = reader.ReadToEnd();
            Assert.Contains("TOC", xml, StringComparison.Ordinal);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordInsertPageBreakAfter_InsertsBreak()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Before break");
            service.WordAppendParagraph(sessionId, "After break");

            var result = service.WordInsertPageBreakAfter(sessionId, 1);
            Assert.Contains("\"ok\":true", result, StringComparison.OrdinalIgnoreCase);

            // Body should now have 3 paragraphs (before, break, after)
            var structure = service.ListStructure(sessionId);
            Assert.Contains("\"bodyParagraphCount\":3", structure, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordSetHeader_AndSetFooter_WritesParts()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAppendParagraph(sessionId, "Body text");

            var hr = service.WordSetHeader(sessionId, "My Document Header");
            var fr = service.WordSetFooter(sessionId, "Page {PAGE} of {NUMPAGES}");

            Assert.Contains("\"ok\":true", hr, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"ok\":true", fr, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);

            // Verify header and footer parts exist in DOCX zip
            using var zipFile = System.IO.Compression.ZipFile.OpenRead(filePath);
            var entries = zipFile.Entries.Select(e => e.FullName).ToList();
            Assert.Contains(entries, e => e.StartsWith("word/header", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(entries, e => e.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordApplyTableStyle_SetsStyleOnTable()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 2, 3);
            var result = service.WordApplyTableStyle(sessionId, 1, "Table Grid");

            Assert.Contains("\"ok\":true", result, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"styleName\":\"Table Grid\"", result, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordFormatTableHeaderRow_AppliesBoldAndShading()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 3, 2);
            service.WordSetTableCell(sessionId, 1, 1, 1, "Header A");
            service.WordSetTableCell(sessionId, 1, 1, 2, "Header B");
            var result = service.WordFormatTableHeaderRow(sessionId, 1, bold: true, shadingFill: "D9EAF7");

            Assert.Contains("\"ok\":true", result, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"tableIndex\":1", result, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordSetTableValues_FillsTableCells()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            service.WordAddTable(sessionId, 3, 3);
            var result = service.WordSetTableValues(
                sessionId, 1,
                "[[\"Component\",\"Location\",\"Status\"],[\"Browser\",\"caQtDM_Web/user\",\"Active\"],[\"Server\",\"caQtDM_Web/srv\",\"Active\"]]");

            Assert.Contains("\"updatedCellCount\":9", result, StringComparison.OrdinalIgnoreCase);

            var cell = service.WordGetTableCell(sessionId, 1, 2, 1);
            Assert.Equal("Browser", cell);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAppendParagraph_ReturnsParagraphIndex()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var r1 = service.WordAppendParagraph(sessionId, "First");
            var r2 = service.WordAppendParagraph(sessionId, "Second");

            using var d1 = JsonDocument.Parse(r1);
            using var d2 = JsonDocument.Parse(r2);

            Assert.Equal(1, d1.RootElement.GetProperty("target").GetProperty("paragraphIndex").GetInt32());
            Assert.Equal(2, d2.RootElement.GetProperty("target").GetProperty("paragraphIndex").GetInt32());

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }

    [Fact]
    public void WordAddHeading_ReturnsParagraphIndex()
    {
        var service = OfficeSessionServiceTestHelpers.CreateService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("docx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "docx");
            var r1 = service.WordAddHeading(sessionId, 1, "Title");
            var r2 = service.WordAppendParagraph(sessionId, "Body");
            var r3 = service.WordAddHeading(sessionId, 2, "Section");

            using var d1 = JsonDocument.Parse(r1);
            using var d3 = JsonDocument.Parse(r3);

            Assert.Equal(1, d1.RootElement.GetProperty("target").GetProperty("paragraphIndex").GetInt32());
            Assert.Equal(3, d3.RootElement.GetProperty("target").GetProperty("paragraphIndex").GetInt32());

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
}
