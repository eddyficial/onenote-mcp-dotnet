// onenote_* tools. Tool names, descriptions, and JSON input schemas are ported
// verbatim from the TypeScript source (onenoteTools.ts); execution goes
// straight to the direct COM layer instead of a PowerShell bridge.

using static OneNoteMcp.Json;

namespace OneNoteMcp;

public static class OneNoteTools
{
    public static List<McpToolDef> CreateAll(OneNoteCom com)
    {
        var tools = new List<McpToolDef>
        {
            // Read
            Hierarchy(com),
            GetPage(com),
            Search(com),
        };
        // Knowledge operator
        tools.AddRange(KnowledgeTools.CreateAll(com));
        tools.AddRange(new[]
        {
            // Page content
            CreatePage(com),
            AppendPage(com),
            UpdatePage(com),
            InsertRichContent(com),
            RenamePage(com),
            MovePage(com),
            ReorderPages(com),
            // Hierarchy
            CreateSection(com),
            CreateSectionGroup(com),
            CreateNotebook(com),
            RenameSection(com),
            MoveSection(com),
            ReorderSections(com),
            // Navigation / export
            Navigate(com),
            Export(com),
            // Delete
            DeletePage(com),
            DeleteSection(com),
            DeleteNotebook(com),
        });
        return tools;
    }

    private static McpToolDef Hierarchy(OneNoteCom com) => new()
    {
        ReadOnly = true,
        Name = "onenote_hierarchy",
        Description =
            "List the OneNote hierarchy: notebooks, section groups, sections, and pages with their IDs. " +
            "Use scope 'notebooks' for a quick overview, 'sections' to find section IDs for page creation, " +
            "'pages' for the full tree including page IDs.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "scope": {
              "type": "string",
              "enum": ["notebooks", "sections", "pages"],
              "description": "How deep to expand the tree. Default 'pages'."
            },
            "start_id": {
              "type": "string",
              "description": "Optional object ID to start from (e.g. one notebook). Empty = all open notebooks."
            }
          },
          "additionalProperties": false
        }
        """,
        Execute = input =>
        {
            var items = com.Hierarchy(input.Str("scope", "pages"), input.Str("start_id"));
            if (items.Count == 0)
            {
                return new ToolResult(
                    "No notebooks are open in OneNote. Ask the user to open a notebook in the OneNote desktop app first.");
            }
            return new ToolResult(AsJson(items));
        },
    };

    private static McpToolDef GetPage(OneNoteCom com) => new()
    {
        ReadOnly = true,
        Name = "onenote_get_page",
        Description =
            "Read a OneNote page: returns its title and text content (flattened from the page XML). " +
            "Get page IDs from onenote_hierarchy or onenote_search.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "page_id": { "type": "string", "description": "OneNote page object ID. Required." }
          },
          "required": ["page_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.GetPage(Require(input, "page_id")))),
    };

    private static McpToolDef Search(OneNoteCom com) => new()
    {
        ReadOnly = true,
        Name = "onenote_search",
        Description =
            "Full-text search across OneNote pages. Returns matching pages with their IDs.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "query": { "type": "string", "description": "Search terms. Required." },
            "start_id": {
              "type": "string",
              "description": "Optional object ID to scope the search (notebook or section)."
            }
          },
          "required": ["query"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.Search(Require(input, "query"), input.Str("start_id")))),
    };

    private static McpToolDef CreatePage(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_create_page",
        Description =
            "Create a new page in a OneNote section, optionally with a title and body text. " +
            "Get the section ID from onenote_hierarchy with scope 'sections'.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "section_id": { "type": "string", "description": "Target section object ID. Required." },
            "title": { "type": "string", "description": "Page title." },
            "body": {
              "type": "string",
              "description": "Initial body text. Newlines become separate paragraphs."
            }
          },
          "required": ["section_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.CreatePage(
            Require(input, "section_id"), input.Str("title"), input.Str("body")))),
    };

    private static McpToolDef AppendPage(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_append_page",
        Description =
            "Append text to the end of an existing OneNote page as a new outline block. " +
            "Newlines become separate paragraphs. Never overwrites existing content.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "page_id": { "type": "string", "description": "Target page object ID. Required." },
            "text": { "type": "string", "description": "Text to append. Required." }
          },
          "required": ["page_id", "text"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.AppendPage(
            Require(input, "page_id"), Require(input, "text")))),
    };

    private static McpToolDef Navigate(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_navigate",
        Description =
            "Open a notebook, section, or page in the visible OneNote window. This changes what the " +
            "user sees on screen — use only when the user asked to open or show something.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "object_id": { "type": "string", "description": "Object ID to navigate to. Required." }
          },
          "required": ["object_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.Navigate(Require(input, "object_id")))),
    };

    private static McpToolDef DeletePage(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Destructive = true,
        Name = "onenote_delete_page",
        Description =
            "Delete a OneNote page by its object ID. By default the page is moved to the notebook's " +
            "recycle bin (recoverable); pass permanent:true to erase it outright. Destructive — always " +
            "confirm you have the right page ID (from onenote_hierarchy or onenote_search) before calling. " +
            "Deletes one page per call.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "page_id": { "type": "string", "description": "OneNote page object ID. Required." },
            "permanent": {
              "type": "boolean",
              "description": "If true, delete permanently instead of moving to the recycle bin. Default false."
            }
          },
          "required": ["page_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.DeleteObject(
            Require(input, "page_id"), input.Bool("permanent")))),
    };

    private static McpToolDef CreateSection(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_create_section",
        Description =
            "Create a new section in a notebook (or inside a section group). Get the notebook/group ID " +
            "from onenote_hierarchy. Returns the new section_id.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "notebook_id": {
              "type": "string",
              "description": "Target notebook (or section group) object ID. Required."
            },
            "section_name": { "type": "string", "description": "Name for the new section. Required." }
          },
          "required": ["notebook_id", "section_name"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.CreateSection(
            Require(input, "notebook_id"), Require(input, "section_name")))),
    };

    private static McpToolDef CreateSectionGroup(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_create_section_group",
        Description =
            "Create a section group inside a notebook or another section group. Returns section_group_id.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "parent_id": {
              "type": "string",
              "description": "Parent notebook or section group object ID. Required."
            },
            "name": { "type": "string", "description": "Name for the new section group. Required." }
          },
          "required": ["parent_id", "name"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.CreateSectionGroup(
            Require(input, "parent_id"), Require(input, "name")))),
    };

    private static McpToolDef CreateNotebook(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_create_notebook",
        Description =
            "Create a new notebook. By default it is created in OneNote's default notebook folder; " +
            "pass an absolute 'path' folder to override. Returns notebook_id.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "name": { "type": "string", "description": "Notebook name. Required." },
            "path": {
              "type": "string",
              "description": "Optional absolute folder to create the notebook in."
            }
          },
          "required": ["name"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.CreateNotebook(
            Require(input, "name"), input.Str("path")))),
    };

    private static McpToolDef RenameSection(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Destructive = true,
        Name = "onenote_rename_section",
        Description = "Rename an existing section.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "section_id": { "type": "string", "description": "Section object ID. Required." },
            "new_name": { "type": "string", "description": "New section name. Required." }
          },
          "required": ["section_id", "new_name"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.RenameSection(
            Require(input, "section_id"), Require(input, "new_name")))),
    };

    private static McpToolDef RenamePage(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Destructive = true,
        Name = "onenote_rename_page",
        Description = "Rename an existing page (sets its title).",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "page_id": { "type": "string", "description": "Page object ID. Required." },
            "new_title": { "type": "string", "description": "New page title. Required." }
          },
          "required": ["page_id", "new_title"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.RenamePage(
            Require(input, "page_id"), Require(input, "new_title")))),
    };

    private static McpToolDef MovePage(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_move_page",
        Description =
            "Move a page to another section. IMPORTANT: OneNote assigns the moved page a NEW object ID — " +
            "use the 'page_id' returned by this tool for any further operations on the page; the old ID " +
            "becomes invalid.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "page_id": { "type": "string", "description": "Page to move. Required." },
            "target_section_id": { "type": "string", "description": "Destination section ID. Required." }
          },
          "required": ["page_id", "target_section_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.MovePage(
            Require(input, "page_id"), Require(input, "target_section_id")))),
    };

    private static McpToolDef MoveSection(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_move_section",
        Description =
            "Move a section into a different notebook or section group (the target parent).",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "section_id": { "type": "string", "description": "Section to move. Required." },
            "target_parent_id": {
              "type": "string",
              "description": "Destination notebook or section group ID. Required."
            }
          },
          "required": ["section_id", "target_parent_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.MoveSection(
            Require(input, "section_id"), Require(input, "target_parent_id")))),
    };

    private static McpToolDef ReorderPages(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_reorder_pages",
        Description =
            "Change page order within a section. Provide either before_page_id or after_page_id as the " +
            "reference the moved page should sit before/after.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "section_id": { "type": "string", "description": "Section holding the pages. Required." },
            "page_id": { "type": "string", "description": "Page to reposition. Required." },
            "before_page_id": { "type": "string", "description": "Place the page before this page." },
            "after_page_id": { "type": "string", "description": "Place the page after this page." }
          },
          "required": ["section_id", "page_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.ReorderPages(
            Require(input, "section_id"), Require(input, "page_id"),
            input.Str("before_page_id"), input.Str("after_page_id")))),
    };

    private static McpToolDef ReorderSections(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_reorder_sections",
        Description =
            "Change section tab order within a notebook or section group. Provide either " +
            "before_section_id or after_section_id as the reference.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "parent_id": {
              "type": "string",
              "description": "Notebook or section group holding the sections. Required."
            },
            "section_id": { "type": "string", "description": "Section to reposition. Required." },
            "before_section_id": { "type": "string", "description": "Place before this section." },
            "after_section_id": { "type": "string", "description": "Place after this section." }
          },
          "required": ["parent_id", "section_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.ReorderSections(
            Require(input, "parent_id"), Require(input, "section_id"),
            input.Str("before_section_id"), input.Str("after_section_id")))),
    };

    private static McpToolDef UpdatePage(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Destructive = true,
        Name = "onenote_update_page",
        Description =
            "Update a page's body content. mode 'replace' (default) clears the existing body outlines and " +
            "writes the new content; mode 'append' adds to the end. The page title is preserved. Newlines " +
            "become separate paragraphs.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "page_id": { "type": "string", "description": "Page object ID. Required." },
            "content": { "type": "string", "description": "New body text. Required." },
            "mode": {
              "type": "string",
              "enum": ["replace", "append"],
              "description": "replace (default) or append."
            }
          },
          "required": ["page_id", "content"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.UpdatePage(
            Require(input, "page_id"), Require(input, "content"), input.Str("mode", "replace")))),
    };

    private static McpToolDef InsertRichContent(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_insert_rich_content",
        Description =
            "Append rich content to a page: a well-formed XHTML fragment and/or an image from a local " +
            "file path. Supported blocks: h1-h6 (rendered as sized bold text), p, ul/ol with nested " +
            "lists, table/tr/th/td, pre (Consolas lines), blockquote, div (recursed). Inline: b/strong, " +
            "i/em, u, code, a href, span with style. Multiple sibling root elements are fine. The " +
            "fragment must parse as XML — self-close void tags and match every open tag. Provide html " +
            "and/or image_path.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "page_id": { "type": "string", "description": "Page object ID. Required." },
            "html": {
              "type": "string",
              "description": "Well-formed XHTML fragment. Block elements become OneNote paragraphs/lists/tables; headings render as sized bold text."
            },
            "image_path": { "type": "string", "description": "Absolute path to an image file to embed." }
          },
          "required": ["page_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.InsertRichContent(
            Require(input, "page_id"), input.Str("html"), input.Str("image_path")))),
    };

    private static McpToolDef Export(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Name = "onenote_export",
        Description =
            "Export a page or section to a file. format: pdf (default), html, docx, mhtml, xps, or onenote. " +
            "target_path is the absolute output file path.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "object_id": { "type": "string", "description": "Page or section object ID. Required." },
            "target_path": { "type": "string", "description": "Absolute output file path. Required." },
            "format": {
              "type": "string",
              "enum": ["pdf", "html", "docx", "mhtml", "xps", "onenote"],
              "description": "Export format. Default pdf."
            }
          },
          "required": ["object_id", "target_path"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.Export(
            Require(input, "object_id"), Require(input, "target_path"), input.Str("format", "pdf")))),
    };

    private static McpToolDef DeleteSection(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Destructive = true,
        Name = "onenote_delete_section",
        Description =
            "Delete a section. Moves it to the notebook's recycle bin by default (recoverable); pass " +
            "permanent:true to erase. Destructive — confirm the section ID first. Deletes the whole " +
            "section and all its pages.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "section_id": { "type": "string", "description": "Section object ID. Required." },
            "permanent": { "type": "boolean", "description": "Erase permanently instead of recycle bin. Default false." }
          },
          "required": ["section_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.DeleteObject(
            Require(input, "section_id"), input.Bool("permanent")))),
    };

    private static McpToolDef DeleteNotebook(OneNoteCom com) => new()
    {
        ReadOnly = false,
        Destructive = true,
        Name = "onenote_delete_notebook",
        Description =
            "Delete/close a whole notebook. Highly destructive — confirm the notebook ID explicitly with " +
            "the user first. Recoverable via recycle bin by default; permanent:true erases.",
        InputSchemaJson = """
        {
          "type": "object",
          "properties": {
            "notebook_id": { "type": "string", "description": "Notebook object ID. Required." },
            "permanent": { "type": "boolean", "description": "Erase permanently. Default false." }
          },
          "required": ["notebook_id"],
          "additionalProperties": false
        }
        """,
        Execute = input => new ToolResult(AsJson(com.DeleteObject(
            Require(input, "notebook_id"), input.Bool("permanent")))),
    };

    internal static string Require(ToolInput input, string key)
    {
        var value = input.Str(key);
        if (value.Length == 0) throw new ArgumentException($"{key} is required");
        return value;
    }
}
