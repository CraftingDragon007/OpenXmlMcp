using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OpenXmlMcp.Server.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXmlMcp.Server.Services;

/// <summary>
/// Handles all PowerPoint (PPTX) document operations. Depends on <see cref="SessionManager"/>
/// for session access and write orchestration.
/// </summary>
public class PowerPointDocumentService(SessionManager sessionManager)
{
    private const int DefaultPptTitleFontSize = 4400;
    private const int DefaultPptBodyFontSize = 2800;
    private const string DefaultPptTitleFont = "Aptos Display";
    private const string DefaultPptBodyFont = "Aptos";

    private enum BodyType
    {
        Text,
        Bulleted,
        Numbered
    }

    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    public string AddSlide(string sessionId, string title, string body, string bodyType = "text")
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(body);
        var parsedBodyType = ParseBodyType(bodyType);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);

        sessionManager.ExecuteWriteOperation(session, "powerpoint_add_slide", () => AddSlideCore(session.FilePath, title, body, parsedBodyType), "power_point_add_slide");
        return SessionManager.BuildMutationResult("powerpoint_add_slide", new { title, bodyType = parsedBodyType.ToString().ToLowerInvariant() }, publicOperationName: "power_point_add_slide", canonicalOperationName: "powerpoint_add_slide");
    }

    public string AddBulletSlide(string sessionId, string title, string bulletLines)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(bulletLines);
        return AddSlide(sessionId, title, bulletLines, "bulleted");
    }

    public string InsertSlideAt(string sessionId, int index, string title, string body)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(body);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);

        sessionManager.ExecuteWriteOperation(session, "powerpoint_insert_slide_at", () =>
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
            newSlidePart.Slide = BuildSlide(title, body, BodyType.Text);
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
        }, "power_point_insert_slide_at");

        return SessionManager.BuildMutationResult("powerpoint_insert_slide_at", new { index, title }, publicOperationName: "power_point_insert_slide_at", canonicalOperationName: "powerpoint_insert_slide_at");
    }

    public string SetSlideTitle(string sessionId, int slideIndex, string title)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        sessionManager.ExecuteWriteOperation(session, "powerpoint_set_slide_title", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
            SetSlideTextBySlot(slidePart, 0, title);
            GetSlide(slidePart).Save();
        }, "power_point_set_slide_title");

        return SessionManager.BuildMutationResult("powerpoint_set_slide_title", new { slideIndex }, publicOperationName: "power_point_set_slide_title", canonicalOperationName: "powerpoint_set_slide_title");
    }

    public string SetSlideBody(string sessionId, int slideIndex, string body, string bodyType = "text")
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(body);
        var parsedBodyType = ParseBodyType(bodyType);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        sessionManager.ExecuteWriteOperation(session, "powerpoint_set_slide_body", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
            SetSlideBodyByType(slidePart, 1, body, parsedBodyType);
            GetSlide(slidePart).Save();
        }, "power_point_set_slide_body");

        return SessionManager.BuildMutationResult("powerpoint_set_slide_body", new { slideIndex, bodyType = parsedBodyType.ToString().ToLowerInvariant() }, publicOperationName: "power_point_set_slide_body", canonicalOperationName: "powerpoint_set_slide_body");
    }

    public string SetSlideNotes(string sessionId, int slideIndex, string notes)
    {
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        sessionManager.ExecuteWriteOperation(session, "powerpoint_set_slide_notes", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
            var notesSlidePart = EnsureNotesSlidePart(slidePart);
            SetNotesSlideText(notesSlidePart, notes ?? string.Empty);
            notesSlidePart.NotesSlide?.Save();
        }, "power_point_set_slide_notes");

        return SessionManager.BuildMutationResult("powerpoint_set_slide_notes", new { slideIndex }, publicOperationName: "power_point_set_slide_notes", canonicalOperationName: "powerpoint_set_slide_notes");
    }

    public string GetSlideNotes(string sessionId, int slideIndex)
    {
        var session = sessionManager.GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.PowerPoint)
        {
            throw new InvalidOperationException("Session document type is not PowerPoint.");
        }

        using var presentation = PresentationDocument.Open(session.FilePath, false);
        var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
        return GetNotesSlideText(slidePart.NotesSlidePart);
    }

    public string SetTextStyle(string sessionId, int slideIndex, int slot, string fontName, int fontSize, bool bold, bool italic, string colorHex)
    {
        SetTextStyleInternal(sessionId, slideIndex, slot, new TextStyle(fontName, fontSize, bold, italic, colorHex));
        return SessionManager.BuildMutationResult("powerpoint_set_text_style", new { slideIndex, slot }, publicOperationName: "power_point_set_text_style", canonicalOperationName: "powerpoint_set_text_style");
    }

    public string GetTextStyle(string sessionId, int slideIndex, int slot)
    {
        var session = sessionManager.GetSession(sessionId);
        if (session.DocumentType != OfficeDocumentType.PowerPoint)
        {
            throw new InvalidOperationException("Session document type is not PowerPoint.");
        }

        using var presentation = PresentationDocument.Open(session.FilePath, false);
        var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
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

        var target = slot switch
        {
            <= 0 => paragraphs[0],
            1 when paragraphs.Count > 1 => paragraphs[1],
            _ => paragraphs[Math.Clamp(slot, 0, paragraphs.Count - 1)]
        };

        A.TextCharacterPropertiesType? runProps = target.Elements<A.Run>().FirstOrDefault()?.RunProperties;
        runProps ??= target.GetFirstChild<A.EndParagraphRunProperties>();

        return JsonSerializer.Serialize(new
        {
            slideIndex,
            slot,
            fontName = runProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value ?? string.Empty,
            fontSize = ((runProps?.FontSize?.Value ?? 0) / 100),
            bold = runProps?.Bold?.Value ?? false,
            italic = runProps?.Italic?.Value ?? false,
            colorHex = runProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value ?? string.Empty
        });
    }

    public string ReorderSlide(string sessionId, int fromIndex, int toIndex)
    {
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        sessionManager.ExecuteWriteOperation(session, "powerpoint_reorder_slide", () =>
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
        }, "power_point_reorder_slide");

        return SessionManager.BuildMutationResult("powerpoint_reorder_slide", new { fromIndex, toIndex }, fromIndex != toIndex, publicOperationName: "power_point_reorder_slide", canonicalOperationName: "powerpoint_reorder_slide");
    }

    public string DeleteSlide(string sessionId, int slideIndex)
    {
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);
        sessionManager.ExecuteWriteOperation(session, "powerpoint_delete_slide", () =>
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
        }, "power_point_delete_slide");

        return SessionManager.BuildMutationResult("powerpoint_delete_slide", new { slideIndex }, publicOperationName: "power_point_delete_slide", canonicalOperationName: "powerpoint_delete_slide");
    }

    // -------------------------------------------------------------------------
    // Internal helpers used by OfficeSessionService for cross-cutting features
    // -------------------------------------------------------------------------

    public void SetTextStyleInternal(string sessionId, int slideIndex, int slot, TextStyle style)
    {
        WordDocumentService.ValidateTextStyle(style);
        var session = sessionManager.GetWritableSession(sessionId, OfficeDocumentType.PowerPoint);

        sessionManager.ExecuteWriteOperation(session, "powerpoint_set_text_style", () =>
        {
            using var presentation = PresentationDocument.Open(session.FilePath, true);
            var slidePart = GetSlidePartByIndex(GetPresentationPart(presentation), slideIndex);
            ApplySlotStyle(slidePart, slot, style);
            GetSlide(slidePart).Save();
        }, "power_point_set_text_style");
    }

    public void ApplyStylePreset(string filePath, string preset)
    {
        using var presentation = PresentationDocument.Open(filePath, true);
        var presentationPart = GetPresentationPart(presentation);
        _ = EnsurePresentationDefaults(presentationPart);

        var slideMasterPart = presentationPart.SlideMasterParts.First();
        var themePart = slideMasterPart.ThemePart ?? slideMasterPart.AddNewPart<ThemePart>();
        WriteDefaultTheme(themePart, ResolveThemeOptions(preset));

        slideMasterPart.SlideMaster?.Save();
        presentationPart.Presentation?.Save();
    }

    public void InitializeEmptyDocument(string filePath)
    {
        using var presentation = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);
        var presentationPart = presentation.AddPresentationPart();
        presentationPart.Presentation = new Presentation();
        _ = EnsurePresentationDefaults(presentationPart);
        presentationPart.Presentation.Save();
    }

    public object ListStructure(string filePath)
    {
        using var presentation = PresentationDocument.Open(filePath, false);
        var presentationPart = GetPresentationPart(presentation);
        var slides = new List<object>();
        var index = 0;
        foreach (var slideId in GetSlideIds(presentationPart))
        {
            index++;
            var slidePart = GetSlidePart(presentationPart, slideId);
            var paragraphs = GetSlide(slidePart)
                .Descendants<P.Shape>()
                .Select(shape => shape.TextBody)
                .Where(textBody => textBody is not null)
                .SelectMany(textBody => textBody!.Elements<A.Paragraph>())
                .Select(p => p.InnerText)
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .ToList();

            slides.Add(new
            {
                slideIndex = index,
                title = paragraphs.ElementAtOrDefault(0) ?? string.Empty,
                body = string.Join("\n", paragraphs.Skip(1)),
                notesPreview = TrimPreview(GetNotesSlideText(slidePart.NotesSlidePart)),
                textSlots = paragraphs
                    .Select((text, slot) => new { slot, text, preview = TrimPreview(text) })
                    .ToArray()
            });
        }

        return new { slideCount = slides.Count, slides };
    }

    public object FindText(string filePath, string query)
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

    public static string[] GetAllowedPresets() => ["default", "neutral"];

    // -------------------------------------------------------------------------
    // Private implementation details
    // -------------------------------------------------------------------------

    private static void AddSlideCore(string filePath, string title, string body, BodyType bodyType)
    {
        using var presentation = PresentationDocument.Open(filePath, true);
        var presentationPart = presentation.PresentationPart ?? throw new InvalidOperationException("Presentation part is missing.");
        var slideLayoutPart = EnsurePresentationDefaults(presentationPart);
        var presentationDocument = presentationPart.Presentation ?? throw new InvalidOperationException("Presentation document is missing.");
        var slideIdList = presentationDocument.SlideIdList ?? presentationDocument.AppendChild(new SlideIdList());

        var newSlidePart = presentationPart.AddNewPart<SlidePart>();
        _ = newSlidePart.AddPart(slideLayoutPart);
        newSlidePart.Slide = BuildSlide(title, body, bodyType);

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
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup()))),
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
            BuildDefaultPresentationTextStyles());

        WriteDefaultTheme(themePart, ResolveThemeOptions("default"));
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
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup(
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

    private static P.Shape CreateLayoutPlaceholderShape(uint id, string name, PlaceholderValues placeholderType, long x, long y, long cx, long cy)
    {
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = placeholderType })),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.EndParagraphRunProperties())));
    }

    private static void WriteDefaultTheme(ThemePart themePart, PresentationThemeOptions options)
    {
        var theme = new A.Theme { Name = options.Name };
        var colorScheme = new A.ColorScheme { Name = "Default" };
        colorScheme.Append(
            new A.Dark1Color(new A.RgbColorModelHex { Val = options.Dark1 }),
            new A.Light1Color(new A.RgbColorModelHex { Val = options.Light1 }),
            new A.Dark2Color(new A.RgbColorModelHex { Val = options.Dark2 }),
            new A.Light2Color(new A.RgbColorModelHex { Val = options.Light2 }),
            new A.Accent1Color(new A.RgbColorModelHex { Val = options.Accent1 }),
            new A.Accent2Color(new A.RgbColorModelHex { Val = options.Accent2 }),
            new A.Accent3Color(new A.RgbColorModelHex { Val = options.Accent3 }),
            new A.Accent4Color(new A.RgbColorModelHex { Val = options.Accent4 }),
            new A.Accent5Color(new A.RgbColorModelHex { Val = options.Accent5 }),
            new A.Accent6Color(new A.RgbColorModelHex { Val = options.Accent6 }),
            new A.Hyperlink(new A.RgbColorModelHex { Val = options.Hyperlink }),
            new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = options.FollowedHyperlink }));

        var fontScheme = new A.FontScheme { Name = options.Name };
        fontScheme.Append(
            new A.MajorFont(new A.LatinFont { Typeface = options.MajorLatinFont }),
            new A.MinorFont(new A.LatinFont { Typeface = options.MinorLatinFont }));

        var formatScheme = new A.FormatScheme { Name = "Office" };
        formatScheme.Append(
            new A.FillStyleList(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
            new A.LineStyleList(
                new A.Outline(new A.PresetDash { Val = A.PresetLineDashValues.Solid }) { Width = 6350 },
                new A.Outline(new A.PresetDash { Val = A.PresetLineDashValues.Solid }) { Width = 12700 },
                new A.Outline(new A.PresetDash { Val = A.PresetLineDashValues.Solid }) { Width = 19050 }),
            new A.EffectStyleList(
                new A.EffectStyle(new A.EffectList()),
                new A.EffectStyle(new A.EffectList()),
                new A.EffectStyle(new A.EffectList())),
            new A.BackgroundFillStyleList(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })));

        theme.Append(new A.ThemeElements(colorScheme, fontScheme, formatScheme));
        themePart.Theme = theme;
        themePart.Theme.Save();
    }

    private static PresentationThemeOptions ResolveThemeOptions(string preset)
    {
        return preset switch
        {
            "neutral" => new PresentationThemeOptions(
                Name: "Neutral",
                Dark1: "1C1C1C",
                Light1: "FFFFFF",
                Dark2: "334155",
                Light2: "E2E8F0",
                Accent1: "2563EB",
                Accent2: "0EA5E9",
                Accent3: "10B981",
                Accent4: "F59E0B",
                Accent5: "EF4444",
                Accent6: "8B5CF6",
                Hyperlink: "1D4ED8",
                FollowedHyperlink: "6D28D9",
                MajorLatinFont: "Calibri",
                MinorLatinFont: "Calibri"),
            _ => new PresentationThemeOptions(
                Name: "Office",
                Dark1: "000000",
                Light1: "FFFFFF",
                Dark2: "1F497D",
                Light2: "EEECE1",
                Accent1: "4F81BD",
                Accent2: "C0504D",
                Accent3: "9BBB59",
                Accent4: "8064A2",
                Accent5: "4BACC6",
                Accent6: "F79646",
                Hyperlink: "0000FF",
                FollowedHyperlink: "800080",
                MajorLatinFont: "Aptos Display",
                MinorLatinFont: "Aptos")
        };
    }

    private static TextStyles BuildDefaultPresentationTextStyles()
    {
        var titleStyle = new TitleStyle();
        titleStyle.Append(new A.Level1ParagraphProperties(new A.DefaultRunProperties { FontSize = DefaultPptTitleFontSize, Language = "en-US" }));
        titleStyle.Append(new A.Level2ParagraphProperties(new A.DefaultRunProperties { FontSize = 3600, Language = "en-US" }));
        titleStyle.Append(new A.Level3ParagraphProperties(new A.DefaultRunProperties { FontSize = 3200, Language = "en-US" }));

        var bodyStyle = new BodyStyle();
        bodyStyle.Append(new A.Level1ParagraphProperties(new A.DefaultRunProperties { FontSize = DefaultPptBodyFontSize, Language = "en-US" }));
        bodyStyle.Append(new A.Level2ParagraphProperties(new A.DefaultRunProperties { FontSize = 2400, Language = "en-US" }));
        bodyStyle.Append(new A.Level3ParagraphProperties(new A.DefaultRunProperties { FontSize = 2000, Language = "en-US" }));

        var otherStyle = new OtherStyle();
        otherStyle.Append(new A.Level1ParagraphProperties(new A.DefaultRunProperties { FontSize = 1800, Language = "en-US" }));
        otherStyle.Append(new A.Level2ParagraphProperties(new A.DefaultRunProperties { FontSize = 1800, Language = "en-US" }));
        otherStyle.Append(new A.Level3ParagraphProperties(new A.DefaultRunProperties { FontSize = 1800, Language = "en-US" }));

        return new TextStyles(titleStyle, bodyStyle, otherStyle);
    }

    private static Slide BuildSlide(string title, string body, BodyType bodyType)
    {
        var shapeTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup()));

        shapeTree.Append(CreateTextShape(2U, "Title", title, 457200L, isTitle: true));
        shapeTree.Append(CreateBodyShape(3U, "Body", body, 1828800L, bodyType));

        return new Slide(new CommonSlideData(shapeTree), new ColorMapOverride(new A.MasterColorMapping()));
    }

    private static P.Shape CreateTextShape(uint id, string name, string text, long yOffset, bool isTitle)
    {
        var fontSize = isTitle ? DefaultPptTitleFontSize : DefaultPptBodyFontSize;
        var fontName = isTitle ? DefaultPptTitleFont : DefaultPptBodyFont;
        var run = new A.Run(new A.Text(text));
        run.PrependChild(CreateDrawingRunProperties(fontSize, fontName));

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 457200L, Y = yOffset },
                    new A.Extents { Cx = 8229600L, Cy = 900000L })),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    run,
                    CreateDrawingEndParagraphRunProperties(fontSize, fontName))));
    }

    private static P.Shape CreateBodyShape(uint id, string name, string text, long yOffset, BodyType bodyType)
    {
        var textBody = new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle());

        foreach (var paragraph in BuildBodyParagraphs(text, bodyType))
        {
            textBody.Append(paragraph);
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 457200L, Y = yOffset },
                    new A.Extents { Cx = 8229600L, Cy = 3600000L })),
            textBody);
    }

    private static List<A.Paragraph> BuildBodyParagraphs(string text, BodyType bodyType)
    {
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var paragraphs = new List<A.Paragraph>();

        if (bodyType == BodyType.Text)
        {
            var run = new A.Run(new A.Text(text));
            run.PrependChild(CreateDrawingRunProperties(DefaultPptBodyFontSize, DefaultPptBodyFont));
            paragraphs.Add(new A.Paragraph(run, CreateDrawingEndParagraphRunProperties(DefaultPptBodyFontSize, DefaultPptBodyFont)));
            return paragraphs;
        }

        foreach (var line in lines)
        {
            var run = new A.Run(new A.Text(line));
            run.PrependChild(CreateDrawingRunProperties(DefaultPptBodyFontSize, DefaultPptBodyFont));

            var bulletElement = bodyType == BodyType.Numbered
                ? (DocumentFormat.OpenXml.OpenXmlElement)new A.AutoNumberedBullet { Type = A.TextAutoNumberSchemeValues.ArabicPeriod, StartAt = 1 }
                : new A.CharacterBullet { Char = "•" };

            var paragraph = new A.Paragraph();
            paragraph.Append(new A.ParagraphProperties(
                bulletElement,
                new A.DefaultRunProperties { FontSize = DefaultPptBodyFontSize, Language = "en-US" })
            {
                Level = 0,
                LeftMargin = 342900,
                Indent = -171450
            });
            paragraph.Append(run);
            paragraph.Append(CreateDrawingEndParagraphRunProperties(DefaultPptBodyFontSize, DefaultPptBodyFont));
            paragraphs.Add(paragraph);
        }

        if (paragraphs.Count == 0)
        {
            paragraphs.Add(new A.Paragraph(CreateDrawingEndParagraphRunProperties(DefaultPptBodyFontSize, DefaultPptBodyFont)));
        }

        return paragraphs;
    }

    private static void ApplySlotStyle(SlidePart slidePart, int slot, TextStyle style)
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

        if (slot <= 0)
        {
            ApplyParagraphStyle(paragraphs[0], style);
            return;
        }

        var bodyParagraphs = paragraphs.Skip(1).ToList();
        if (bodyParagraphs.Count == 0)
        {
            ApplyParagraphStyle(paragraphs[0], style);
            return;
        }

        if (slot == 1)
        {
            foreach (var paragraph in bodyParagraphs)
            {
                ApplyParagraphStyle(paragraph, style);
            }

            return;
        }

        var targetIndex = Math.Clamp(slot, 0, paragraphs.Count - 1);
        ApplyParagraphStyle(paragraphs[targetIndex], style);
    }

    private static void ApplyParagraphStyle(A.Paragraph paragraph, TextStyle style)
    {
        foreach (var run in paragraph.Elements<A.Run>())
        {
            run.RunProperties ??= new A.RunProperties();
            ApplyDrawingRunProperties(run.RunProperties, style);
        }

        var endParagraphProperties = paragraph.GetFirstChild<A.EndParagraphRunProperties>() ?? paragraph.AppendChild(new A.EndParagraphRunProperties());
        ApplyDrawingRunProperties(endParagraphProperties, style);
    }

    private static void ApplyDrawingRunProperties(A.TextCharacterPropertiesType runProperties, TextStyle style)
    {
        runProperties.FontSize = style.FontSize * 100;
        runProperties.Bold = style.Bold;
        runProperties.Italic = style.Italic;
        runProperties.RemoveAllChildren<A.LatinFont>();
        runProperties.Append(new A.LatinFont { Typeface = style.FontName });
        runProperties.RemoveAllChildren<A.SolidFill>();
        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = style.ColorHex }));
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

    private static void SetSlideBodyByType(SlidePart slidePart, int slot, string body, BodyType bodyType)
    {
        var textBodies = GetSlide(slidePart)
            .Descendants<P.Shape>()
            .Select(shape => shape.TextBody)
            .Where(textBody => textBody is not null)
            .ToList();

        if (textBodies.Count == 0)
        {
            throw new InvalidOperationException("Slide does not contain editable text placeholders.");
        }

        var targetIndex = Math.Clamp(slot, 0, textBodies.Count - 1);
        var textBody = textBodies[targetIndex]!;
        textBody.RemoveAllChildren<A.Paragraph>();

        foreach (var paragraph in BuildBodyParagraphs(body, bodyType))
        {
            textBody.Append(paragraph);
        }
    }

    private static NotesSlidePart EnsureNotesSlidePart(SlidePart slidePart)
    {
        if (slidePart.NotesSlidePart is not null)
        {
            slidePart.NotesSlidePart.NotesSlide ??= BuildDefaultNotesSlide();
            return slidePart.NotesSlidePart;
        }

        var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
        notesSlidePart.NotesSlide = BuildDefaultNotesSlide();
        return notesSlidePart;
    }

    private static NotesSlide BuildDefaultNotesSlide()
    {
        var shapeTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup()));

        shapeTree.Append(new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2U, Name = "Notes Placeholder" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1U })),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 457200L, Y = 914400L },
                    new A.Extents { Cx = 8229600L, Cy = 4572000L }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.EndParagraphRunProperties()))));

        return new NotesSlide(new CommonSlideData(shapeTree), new ColorMapOverride(new A.MasterColorMapping()));
    }

    private static void SetNotesSlideText(NotesSlidePart notesSlidePart, string text)
    {
        var paragraph = notesSlidePart.NotesSlide?
            .Descendants<P.Shape>()
            .Select(shape => shape.TextBody)
            .Where(body => body is not null)
            .SelectMany(body => body!.Elements<A.Paragraph>())
            .FirstOrDefault();

        if (paragraph is null)
        {
            throw new InvalidOperationException("Notes slide does not contain editable text placeholders.");
        }

        paragraph.RemoveAllChildren<A.Run>();
        paragraph.RemoveAllChildren<A.Break>();
        paragraph.RemoveAllChildren<A.Field>();
        paragraph.Append(new A.Run(new A.Text(text)));
    }

    private static string GetNotesSlideText(NotesSlidePart? notesSlidePart)
    {
        if (notesSlidePart?.NotesSlide is null)
        {
            return string.Empty;
        }

        return string.Join("\n", notesSlidePart.NotesSlide
            .Descendants<P.Shape>()
            .Select(shape => shape.TextBody)
            .Where(body => body is not null)
            .SelectMany(body => body!.Elements<A.Paragraph>())
            .Select(p => p.InnerText)
            .Where(t => !string.IsNullOrWhiteSpace(t)));
    }

    private static A.RunProperties CreateDrawingRunProperties(int fontSize, string fontName)
    {
        var properties = new A.RunProperties { FontSize = fontSize, Language = "en-US" };
        properties.Append(new A.LatinFont { Typeface = fontName });
        return properties;
    }

    private static A.EndParagraphRunProperties CreateDrawingEndParagraphRunProperties(int fontSize, string fontName)
    {
        var properties = new A.EndParagraphRunProperties { FontSize = fontSize, Language = "en-US" };
        properties.Append(new A.LatinFont { Typeface = fontName });
        return properties;
    }

    private static SlideIdList GetOrCreateSlideIdList(PresentationPart presentationPart)
    {
        var presentation = presentationPart.Presentation ?? throw new InvalidOperationException("Presentation document is missing.");
        return presentation.SlideIdList ?? presentation.AppendChild(new SlideIdList());
    }

    private static uint GetNextSlideId(SlideIdList slideIdList)
        => slideIdList.Elements<SlideId>().Select(x => x.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max() + 1;

    private static SlidePart GetSlidePartByIndex(PresentationPart presentationPart, int slideIndex)
    {
        var slideIds = GetSlideIds(presentationPart).ToList();
        if (slideIndex < 1 || slideIndex > slideIds.Count)
        {
            throw new InvalidOperationException($"Slide index {slideIndex} is out of range. Valid range is 1..{slideIds.Count}.");
        }

        return GetSlidePart(presentationPart, slideIds[slideIndex - 1]);
    }

    private static PresentationPart GetPresentationPart(PresentationDocument presentation)
        => presentation.PresentationPart ?? throw new InvalidOperationException("Presentation part is missing.");

    private static IEnumerable<SlideId> GetSlideIds(PresentationPart presentationPart)
        => presentationPart.Presentation?.SlideIdList?.Elements<SlideId>() ?? [];

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
        => slidePart.Slide ?? throw new InvalidOperationException("Slide content is missing.");

    private static BodyType ParseBodyType(string bodyType)
    {
        var normalized = string.IsNullOrWhiteSpace(bodyType) ? "text" : bodyType.Trim().ToLowerInvariant();
        return normalized switch
        {
            "text" => BodyType.Text,
            "bulleted" => BodyType.Bulleted,
            "numbered" => BodyType.Numbered,
            _ => throw new InvalidOperationException("Invalid bodyType. Allowed values: text, bulleted, numbered.")
        };
    }

    private static string TrimPreview(string text, int maxLength = 120)
    {
        var normalized = string.Join(" ", (text ?? string.Empty).Split('\n', '\r').Select(x => x.Trim()).Where(x => x.Length > 0));
        if (normalized.Length <= maxLength) return normalized;
        return normalized[..maxLength] + "...";
    }
}
