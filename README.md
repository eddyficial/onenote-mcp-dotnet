# OneNote MCP for Windows (.NET)

A standalone Model Context Protocol server for Microsoft OneNote desktop,
written in C# on .NET 8. It connects to OneNote through its Windows COM API
**directly from .NET** — no PowerShell bridge process, no API key, no embedded
AI provider.

This is a full port of the Bun/TypeScript `onenote-mcp-windows` server: the
same 27 `onenote_*` tools, with identical tool names, input schemas, and
behavior.

## Why no PowerShell bridge?

The TypeScript version needed a persistent PowerShell child process because
Node cannot speak COM. .NET can, so this server talks to OneNote in-process.

One subtlety: with x64 Click-to-Run Office, the OneNote typelib is registered
only under the Win32 registry key, which breaks every IDispatch late-binding
path from a 64-bit process (pywin32, C# `dynamic`, `Type.InvokeMember` all
fail with `TYPE_E_LIBNOTREGISTERED` / `E_FAIL`). This server therefore calls
OneNote through the `IApplication` dual-interface **vtable** (early binding
against the documented interface layout), which bypasses typelib lookups
entirely. All COM calls run on a single dedicated STA thread.

## Requirements

- Windows
- Microsoft OneNote desktop from Office (Click-to-Run or MSI), with at least
  one notebook open
- .NET 8 SDK or newer (any SDK that can target `net8.0-windows`)

## Build

```powershell
git clone <this repo>
cd onenote-mcp-dotnet
dotnet build
```

## Run

```powershell
dotnet run --project OneNoteMcp
```

Or run the built executable directly:

```powershell
.\OneNoteMcp\bin\Debug\net8.0-windows\OneNoteMcp.exe
```

The server speaks MCP over stdio: stdout carries JSON-RPC only; all
diagnostics go to stderr.

## Connect a client

### Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "dotnet",
      "args": ["run", "--project", "C:\\path\\to\\onenote-mcp-dotnet\\OneNoteMcp"]
    }
  }
}
```

Or point at the built exe (faster startup, no build check):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "C:\\path\\to\\onenote-mcp-dotnet\\OneNoteMcp\\bin\\Debug\\net8.0-windows\\OneNoteMcp.exe"
    }
  }
}
```

### Claude Code

```powershell
claude mcp add onenote -- dotnet run --project C:\path\to\onenote-mcp-dotnet\OneNoteMcp
```

### Codex

`%USERPROFILE%\.codex\config.toml`:

```toml
[mcp_servers.onenote]
command = "dotnet"
args = ["run", "--project", "C:\\path\\to\\onenote-mcp-dotnet\\OneNoteMcp"]
```

## Tools

| Tool | Kind | Description |
|---|---|---|
| `onenote_hierarchy` | read | List notebooks / section groups / sections / pages with IDs |
| `onenote_get_page` | read | Read a page's title and flattened text |
| `onenote_search` | read | Full-text search across pages |
| `onenote_knowledge_digest` | read | Source-grounded digest across a scope |
| `onenote_extract_insights` | read | Action items, owners, due dates, decisions, risks, questions |
| `onenote_health_report` | read | Duplicates, stale/untitled/empty pages, org score |
| `onenote_template_preview` | read | Preview a trusted template (meeting, project, decision_log, weekly_review) |
| `onenote_create_from_template` | write | Create a page from a template (preview-first) |
| `onenote_weekly_review` | read | Weekly-review source pack |
| `onenote_create_page` | write | Create a page in a section |
| `onenote_append_page` | write | Append text to a page |
| `onenote_update_page` | write | Replace or append page body |
| `onenote_insert_rich_content` | write | Append XHTML fragment and/or image |
| `onenote_rename_page` | write | Set a page title |
| `onenote_move_page` | write | Move a page to another section (returns the new page ID) |
| `onenote_reorder_pages` | write | Reorder pages within a section |
| `onenote_create_section` | write | Create a section |
| `onenote_create_section_group` | write | Create a section group |
| `onenote_create_notebook` | write | Create a notebook |
| `onenote_rename_section` | write | Rename a section |
| `onenote_move_section` | write | Move a section to another parent |
| `onenote_reorder_sections` | write | Reorder section tabs |
| `onenote_navigate` | write | Open an object in the OneNote window |
| `onenote_export` | write | Export page/section to pdf, html, docx, mhtml, xps, or onenote |
| `onenote_delete_page` | write | Delete a page (recycle bin by default) |
| `onenote_delete_section` | write | Delete a section (recycle bin by default) |
| `onenote_delete_notebook` | write | Delete/close a notebook (recycle bin by default) |

## Safety

- Template creation previews by default (`preview_only` defaults to true).
- Delete tools use OneNote's recycle bin unless `permanent: true` is explicit.
- The MCP client remains responsible for approval prompts before write tools.

## Test

```powershell
dotnet test
```

## License

MIT
