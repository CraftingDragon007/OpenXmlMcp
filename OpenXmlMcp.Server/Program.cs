using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OpenXmlMcp.Server.Services;
using OpenXmlMcp.Server.Tools;

var builder = Host.CreateApplicationBuilder(args);

// Configure all logs to go to stderr (stdout is used for the MCP protocol messages).
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);

// Add the MCP services: the transport to use (stdio) and the tools to register.
builder.Services
    .AddSingleton<OpenXmlDocumentService>()
    .AddSingleton<SessionManager>()
    .AddSingleton<WordDocumentService>()
    .AddSingleton<ExcelDocumentService>()
    .AddSingleton<PowerPointDocumentService>()
    .AddSingleton<OfficeSessionService>()
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithTools<OpenXmlTools>()
    .WithTools<OfficeTools>();

await builder.Build().RunAsync();
