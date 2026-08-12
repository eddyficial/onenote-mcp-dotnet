// Standalone OneNote MCP server for Windows (.NET port).
//
// Exposes the same onenote_* tools as the TypeScript original, but talks to
// OneNote through direct late-bound COM — no PowerShell bridge process.
// Tools are hand-registered with explicit JSON schemas so tool names and input
// schemas match the TypeScript source verbatim.
//
// stdio transport uses stdout for the JSON-RPC protocol, so all diagnostics go
// to stderr only.

using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using OneNoteMcp;

using var com = new OneNoteCom();
var tools = OneNoteTools.CreateAll(com);
var byName = tools.ToDictionary(t => t.Name);

var listedTools = tools.Select(t => new Tool
{
    Name = t.Name,
    Description = t.Description,
    InputSchema = JsonSerializer.Deserialize<JsonElement>(t.InputSchemaJson),
    Annotations = new ToolAnnotations
    {
        ReadOnlyHint = t.ReadOnly,
        DestructiveHint = !t.ReadOnly && t.Destructive,
        // Everything talks to the local OneNote instance only.
        OpenWorldHint = false,
    },
}).ToList();

var options = new McpServerOptions
{
    ServerInfo = new Implementation { Name = "onenote-mcp-dotnet", Version = "0.1.2" },
    Handlers = new McpServerHandlers
    {
        ListToolsHandler = (request, cancellationToken) =>
            ValueTask.FromResult(new ListToolsResult { Tools = listedTools }),
        CallToolHandler = async (request, cancellationToken) =>
        {
            var name = request.Params?.Name ?? "";
            if (!byName.TryGetValue(name, out var tool))
            {
                return new CallToolResult
                {
                    Content = [new TextContentBlock { Text = $"Unknown tool: {name}" }],
                    IsError = true,
                };
            }
            try
            {
                // Task.Run keeps the protocol loop responsive; the COM layer's
                // single STA work queue still serializes OneNote access.
                var input = new ToolInput(request.Params?.Arguments);
                var result = await Task.Run(() => tool.Execute(input), cancellationToken);
                return new CallToolResult
                {
                    Content = [new TextContentBlock { Text = result.Output }],
                    IsError = result.IsError,
                };
            }
            catch (Exception ex)
            {
                return new CallToolResult
                {
                    Content = [new TextContentBlock { Text = ex.Message }],
                    IsError = true,
                };
            }
        },
    },
};

await using var server = McpServer.Create(new StdioServerTransport("onenote-mcp-dotnet"), options);
Console.Error.WriteLine("onenote-mcp-dotnet ready (stdio)");
await server.RunAsync();
