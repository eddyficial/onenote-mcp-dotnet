// Shared tool model: name + description + verbatim JSON input schema + an
// execute function. Mirrors the McpTool shape of the TypeScript source so the
// MCP layer stays a thin adapter.

using System.Text.Encodings.Web;
using System.Text.Json;

namespace OneNoteMcp;

public sealed record ToolResult(string Output, bool IsError = false);

public sealed class McpToolDef
{
    public required bool ReadOnly { get; init; }
    // True for tools that destroy or overwrite existing user content (deletes,
    // replace-mode updates, renames). Surfaced as the MCP destructiveHint so
    // clients can permission-gate these more strictly than additive writes.
    public bool Destructive { get; init; }
    public required string Name { get; init; }
    public required string Description { get; init; }
    public required string InputSchemaJson { get; init; }
    public required Func<ToolInput, ToolResult> Execute { get; init; }
}

public sealed class ToolInput
{
    private readonly IDictionary<string, JsonElement>? _args;

    public ToolInput(IDictionary<string, JsonElement>? args) => _args = args;

    public string Str(string key, string fallback = "")
    {
        if (_args is null || !_args.TryGetValue(key, out var value)) return fallback;
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString() ?? fallback,
            JsonValueKind.Null or JsonValueKind.Undefined => fallback,
            _ => value.ToString(),
        };
    }

    public bool Bool(string key, bool fallback = false)
    {
        if (_args is null || !_args.TryGetValue(key, out var value)) return fallback;
        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            _ => fallback,
        };
    }

    public int Int(string key, int fallback)
    {
        if (_args is null || !_args.TryGetValue(key, out var value)) return fallback;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var d) && d != 0)
            return (int)d;
        return fallback;
    }

    public bool IsExplicitFalse(string key) =>
        _args is not null && _args.TryGetValue(key, out var value) && value.ValueKind == JsonValueKind.False;
}

public static class Json
{
    // Matches JSON.stringify(value, null, 2): two-space indent, minimal escaping.
    public static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static string AsJson(object? value) => JsonSerializer.Serialize(value, Options);
}
