// Knowledge-operator tools built on top of the read ops: digest, insights,
// health report, templates, weekly review. Port of knowledgeTools.ts.

using static OneNoteMcp.Json;

namespace OneNoteMcp;

public static class KnowledgeTools
{
    private static List<PageRecord> LoadPages(OneNoteCom com, string startId = "", int maxPages = 100)
    {
        var hierarchy = com.Hierarchy("pages", startId);
        var pageNodes = KnowledgeOperator.FlattenPages(hierarchy)
            .Take(Math.Max(1, Math.Min(maxPages, 500)))
            .ToList();
        var pages = new List<PageRecord>();
        foreach (var node in pageNodes)
        {
            var page = com.GetPage(node.Id);
            pages.Add(new PageRecord
            {
                PageId = page.PageId,
                Title = page.Title.Length > 0 ? page.Title : (node.Name.Length > 0 ? node.Name : "Untitled"),
                Text = page.Text,
                Path = node.Path,
                LastModifiedTime = node.LastModifiedTime,
            });
        }
        return pages;
    }

    private const string ScopeProperties = """
            "start_id": { "type": "string", "description": "Optional notebook, section group, or section ID. Empty means all open notebooks." },
            "max_pages": { "type": "number", "description": "Safety cap from 1 to 500. Default 100." }
        """;

    public static List<McpToolDef> CreateAll(OneNoteCom com) => new()
    {
        new McpToolDef
        {
            ReadOnly = true,
            Name = "onenote_knowledge_digest",
            Description = "Build a source-grounded executive or detailed digest across a page, section, notebook, or all open notebooks. Returns per-page key points plus action items, decisions, risks, and questions for the client to synthesize.",
            InputSchemaJson = $$"""
            {
              "type": "object",
              "properties": {
                {{ScopeProperties}},
                "mode": { "type": "string", "enum": ["executive", "detailed"] }
              },
              "additionalProperties": false
            }
            """,
            Execute = input => new ToolResult(AsJson(KnowledgeOperator.BuildKnowledgeDigest(
                LoadPages(com, input.Str("start_id"), input.Int("max_pages", 100)),
                input.Str("mode") == "detailed" ? "detailed" : "executive"))),
        },
        new McpToolDef
        {
            ReadOnly = true,
            Name = "onenote_extract_insights",
            Description = "Extract action items, owners, due dates, decisions, risks, blockers, and open questions from a scoped set of OneNote pages.",
            InputSchemaJson = $$"""
            {
              "type": "object",
              "properties": {
                {{ScopeProperties}}
              },
              "additionalProperties": false
            }
            """,
            Execute = input => new ToolResult(AsJson(KnowledgeOperator.ExtractInsights(
                LoadPages(com, input.Str("start_id"), input.Int("max_pages", 100))))),
        },
        new McpToolDef
        {
            ReadOnly = true,
            Name = "onenote_health_report",
            Description = "Audit a notebook or section for duplicate candidates, stale pages, untitled/empty pages, ownerless action items, and organization recommendations with explainable confidence scores.",
            InputSchemaJson = $$"""
            {
              "type": "object",
              "properties": {
                {{ScopeProperties}},
                "stale_days": { "type": "number", "description": "Age threshold in days. Default 180." }
              },
              "additionalProperties": false
            }
            """,
            Execute = input => new ToolResult(AsJson(KnowledgeOperator.BuildHealthReport(
                LoadPages(com, input.Str("start_id"), input.Int("max_pages", 100)),
                DateTimeOffset.UtcNow,
                input.Int("stale_days", 180)))),
        },
        new McpToolDef
        {
            ReadOnly = true,
            Name = "onenote_template_preview",
            Description = "Preview a trusted page template without modifying OneNote. Available templates: meeting, project, decision_log, weekly_review.",
            InputSchemaJson = """
            {
              "type": "object",
              "properties": {
                "template": { "type": "string", "enum": ["meeting", "project", "decision_log", "weekly_review"] },
                "title": { "type": "string" },
                "date": { "type": "string" }
              },
              "required": ["template", "title"],
              "additionalProperties": false
            }
            """,
            Execute = input =>
            {
                var template = input.Str("template");
                var title = input.Str("title");
                var date = input.Str("date");
                return new ToolResult(AsJson(new Dictionary<string, object?>
                {
                    ["preview"] = true,
                    ["template"] = template,
                    ["title"] = title,
                    ["body"] = KnowledgeOperator.RenderTemplate(template, title, date.Length > 0 ? date : null),
                    ["explanation"] = "No OneNote content was changed.",
                }));
            },
        },
        new McpToolDef
        {
            ReadOnly = false,
            Name = "onenote_create_from_template",
            Description = "Create a page from a trusted template after preview. preview_only defaults true; set false only after the user approves the rendered content.",
            InputSchemaJson = """
            {
              "type": "object",
              "properties": {
                "section_id": { "type": "string" },
                "template": { "type": "string", "enum": ["meeting", "project", "decision_log", "weekly_review"] },
                "title": { "type": "string" },
                "date": { "type": "string" },
                "preview_only": { "type": "boolean", "description": "Default true." }
              },
              "required": ["section_id", "template", "title"],
              "additionalProperties": false
            }
            """,
            Execute = input =>
            {
                var template = input.Str("template");
                var title = input.Str("title");
                var date = input.Str("date");
                var sectionId = input.Str("section_id");
                var body = KnowledgeOperator.RenderTemplate(template, title, date.Length > 0 ? date : null);
                if (!input.IsExplicitFalse("preview_only"))
                {
                    return new ToolResult(AsJson(new Dictionary<string, object?>
                    {
                        ["preview"] = true,
                        ["would_create_in"] = sectionId,
                        ["title"] = title,
                        ["body"] = body,
                        ["explanation"] = "Set preview_only=false after approval to create the page.",
                    }));
                }
                var created = com.CreatePage(sectionId, title, body);
                return new ToolResult(AsJson(new Dictionary<string, object?>
                {
                    ["preview"] = false,
                    ["created"] = created,
                    ["template"] = template,
                    ["title"] = title,
                    ["change_log"] = new Dictionary<string, object?>
                    {
                        ["operation"] = "create_page_from_template",
                        ["timestamp"] = DateTimeOffset.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'"),
                    },
                }));
            },
        },
        new McpToolDef
        {
            ReadOnly = true,
            Name = "onenote_weekly_review",
            Description = "Build a weekly-review source pack from recent/scoped notes, including key points, decisions, risks, and incomplete action items. This is read-only and does not create a page.",
            InputSchemaJson = $$"""
            {
              "type": "object",
              "properties": {
                {{ScopeProperties}}
              },
              "additionalProperties": false
            }
            """,
            Execute = input =>
            {
                var pages = LoadPages(com, input.Str("start_id"), input.Int("max_pages", 100));
                return new ToolResult(AsJson(new Dictionary<string, object?>
                {
                    ["template"] = KnowledgeOperator.RenderTemplate("weekly_review", "Weekly Review"),
                    ["source_pack"] = KnowledgeOperator.BuildKnowledgeDigest(pages, "detailed"),
                    ["explanation"] = "Review and edit this source pack before creating a page.",
                }));
            },
        },
    };
}
