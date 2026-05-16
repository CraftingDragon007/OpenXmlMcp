using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlMcp.Server.Services;

namespace OpenXmlMcp.Server.Tests;

public partial class OfficeSessionServiceTests
{
    [Fact]
    public void PowerPointSession_AddSlideAndFind_Works()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Roadmap", "Phase one complete");

            var structure = service.ListStructure(sessionId);
            Assert.Contains("slideCount", structure);

            var findResult = service.FindText(sessionId, "phase one");
            Assert.Contains("matchCount", findResult);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointSession_AddBulletSlide_IsSearchable()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddBulletSlide(sessionId, "Sprint", "Task A\nTask B");

            var result = service.FindText(sessionId, "Task B");
            Assert.Contains("\"matchCount\":1", result, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointStructure_ContainsSlidesMetadata()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Intro", "Body");

            var structure = service.ListStructure(sessionId);
            Assert.Contains("\"slides\"", structure, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"slideIndex\":1", structure, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"title\":\"Intro\"", structure, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointInsertSlideAt_UsesOneBasedIndex()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "One", "First");
            service.PowerPointAddSlide(sessionId, "Three", "Third");
            service.PowerPointInsertSlideAt(sessionId, 2, "Two", "Second");

            var search = service.FindText(sessionId, "Two");
            Assert.Contains("\"slideIndex\":2", search, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointSetSlideTitleAndBody_UpdatesSlideContent()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Old Title", "Old Body");
            service.PowerPointSetSlideTitle(sessionId, 1, "New Title");
            service.PowerPointSetSlideBody(sessionId, 1, "New Body");

            var titleSearch = service.FindText(sessionId, "New Title");
            var bodySearch = service.FindText(sessionId, "New Body");
            Assert.Contains("\"matchCount\":1", titleSearch, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"matchCount\":1", bodySearch, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointSetSlideBody_CanConvertTextToBulletedOrNumbered()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Old Title", "Single line");
            service.PowerPointSetSlideBody(sessionId, 1, "One\nTwo", "bulleted");
            service.PowerPointSetSlideBody(sessionId, 1, "First\nSecond", "numbered");
            service.CloseDocument(sessionId);

            using var presentation = PresentationDocument.Open(filePath, false);
            var slidePart = presentation.PresentationPart!.SlideParts.First();
            Assert.NotNull(slidePart.Slide);
            var autoNumbered = slidePart.Slide!
                .Descendants<DocumentFormat.OpenXml.Drawing.ParagraphProperties>()
                .Count(p => p.Descendants<DocumentFormat.OpenXml.Drawing.AutoNumberedBullet>().Any());

            Assert.True(autoNumbered >= 2);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointStructure_BodyContainsAllListLines()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Plan", "Kickoff\nScope\nDelivery", "bulleted");

            var structure = service.ListStructure(sessionId);
            Assert.Contains("Kickoff", structure, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Scope", structure, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Delivery", structure, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointSetAndGetSlideNotes_Works()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Topic", "Visible body");
            service.PowerPointSetSlideNotes(sessionId, 1, "Hidden presenter note");

            var notes = service.PowerPointGetSlideNotes(sessionId, 1);
            Assert.Equal("Hidden presenter note", notes);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointStructure_IncludesNotesPreview()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Topic", "Visible body");
            service.PowerPointSetSlideNotes(sessionId, 1, "Presenter-only guidance");

            var structure = service.ListStructure(sessionId);
            Assert.Contains("\"notesPreview\":\"Presenter-only guidance\"", structure, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointReorderAndDeleteSlide_Work()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "One", "A");
            service.PowerPointAddSlide(sessionId, "Two", "B");
            service.PowerPointAddSlide(sessionId, "Three", "C");

            service.PowerPointReorderSlide(sessionId, 3, 1);
            var movedSearch = service.FindText(sessionId, "Three");
            Assert.Contains("\"slideIndex\":1", movedSearch, StringComparison.OrdinalIgnoreCase);

            service.PowerPointDeleteSlide(sessionId, 2);
            var structure = service.ListStructure(sessionId);
            Assert.Contains("\"slideCount\":2", structure, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointCreateDocument_HasSlideMasterAndTheme()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.CloseDocument(sessionId);

            using var presentation = PresentationDocument.Open(filePath, false);
            var presentationPart = presentation.PresentationPart;

            Assert.NotNull(presentationPart);
            Assert.NotEmpty(presentationPart!.SlideMasterParts);
            var slideMasterPart = presentationPart.SlideMasterParts.First();
            Assert.NotNull(slideMasterPart.ThemePart);
            Assert.NotEmpty(slideMasterPart.SlideLayoutParts);
            Assert.Equal("Office", slideMasterPart.ThemePart!.Theme?.Name?.Value);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointAddSlide_UsesReadableDefaultFontSizes()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Title", "Body");
            service.CloseDocument(sessionId);

            using var presentation = PresentationDocument.Open(filePath, false);
            var slidePart = presentation.PresentationPart!.SlideParts.First();
            Assert.NotNull(slidePart.Slide);
            var paragraphs = slidePart.Slide!.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().ToList();
            Assert.True(paragraphs.Count >= 2);

            var titleRunProps = paragraphs[0].Descendants<DocumentFormat.OpenXml.Drawing.RunProperties>().FirstOrDefault();
            var bodyRunProps = paragraphs[1].Descendants<DocumentFormat.OpenXml.Drawing.RunProperties>().FirstOrDefault();

            Assert.Equal(4400, titleRunProps?.FontSize?.Value);
            Assert.Equal(2800, bodyRunProps?.FontSize?.Value);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointSetTextStyle_AndApplyTextPreset_Work()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Title", "Body");
            service.PowerPointSetTextStyle(sessionId, 1, 0, "Calibri", 30, bold: true, italic: false, colorHex: "AA0000");
            service.ApplyTextPreset(sessionId, "subtitle", 1);

            var payload = service.ValidateOperation(sessionId, "apply_text_preset");
            Assert.Contains("\"isValid\":true", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
    [Fact]
    public void PowerPointGetTextStyle_ReturnsFontFields()
    {
        var service = new OfficeSessionService();
        var filePath = OfficeSessionServiceTestHelpers.GetTempPath("pptx");

        try
        {
            var sessionId = service.CreateDocument(filePath, "pptx");
            service.PowerPointAddSlide(sessionId, "Title", "Body");
            service.PowerPointSetTextStyle(sessionId, 1, 0, "Calibri", 28, true, false, "112233");

            var payload = service.PowerPointGetTextStyle(sessionId, 1, 0);
            Assert.Contains("\"slideIndex\":1", payload, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("\"fontSize\":28", payload, StringComparison.OrdinalIgnoreCase);

            service.CloseDocument(sessionId);
        }
        finally
        {
            OfficeSessionServiceTestHelpers.DeleteIfExists(filePath);
        }
    }
}
