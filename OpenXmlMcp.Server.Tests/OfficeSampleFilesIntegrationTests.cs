using OpenXmlMcp.Server.Services;

namespace OpenXmlMcp.Server.Tests;

public class OfficeSampleFilesIntegrationTests
{
    private readonly OfficeSessionService _service = OfficeSessionServiceTestHelpers.CreateService();

    [Fact]
    public void SampleDocxFile_CanBeOpenedAndInspected()
    {
        var path = ResolveRootFile("file-sample_100kB.docx");
        var sessionId = _service.OpenDocument(path, readOnly: true);

        try
        {
            var structure = _service.ListStructure(sessionId);
            var search = _service.FindText(sessionId, "lorem");

            Assert.Contains("paragraphCount", structure);
            Assert.Contains("matchCount", search);
        }
        finally
        {
            _service.CloseDocument(sessionId);
        }
    }

    [Fact]
    public void SampleXlsxFile_CanBeOpenedAndInspected()
    {
        var path = ResolveRootFile("file_example_XLSX_1000.xlsx");
        var sessionId = _service.OpenDocument(path, readOnly: true);

        try
        {
            var structure = _service.ListStructure(sessionId);

            Assert.Contains("sheetCount", structure);
            Assert.Contains("Sheet", structure, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            _service.CloseDocument(sessionId);
        }
    }

    [Fact]
    public void SamplePptxFile_CanBeOpenedAndInspected()
    {
        var path = ResolveRootFile("file_example_PPTX_250kB.pptx");
        var sessionId = _service.OpenDocument(path, readOnly: true);

        try
        {
            var structure = _service.ListStructure(sessionId);
            var search = _service.FindText(sessionId, "lorem");

            Assert.Contains("slideCount", structure);
            Assert.Contains("matchCount", search);
        }
        finally
        {
            _service.CloseDocument(sessionId);
        }
    }

    private static string ResolveRootFile(string fileName)
    {
        var rootPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var path = Path.Combine(rootPath, fileName);
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Sample file not found: {path}");
        }

        return path;
    }
}
