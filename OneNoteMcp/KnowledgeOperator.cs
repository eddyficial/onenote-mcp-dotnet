// Deterministic knowledge-operator core: insight extraction, health report,
// digest, similarity, and trusted templates. Pure functions — no COM — so the
// whole file is unit-testable. Port of the TypeScript knowledgeOperator.ts.

using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace OneNoteMcp;

public sealed class HierarchyNode
{
    [JsonPropertyName("kind")] public string Kind { get; set; } = "";
    [JsonPropertyName("id")] public string Id { get; set; } = "";
    [JsonPropertyName("name")] public string Name { get; set; } = "";
    [JsonPropertyName("currentlyViewed")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? CurrentlyViewed { get; set; }
    [JsonPropertyName("lastModifiedTime")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LastModifiedTime { get; set; }
    [JsonPropertyName("dateTime")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DateTime { get; set; }
    [JsonPropertyName("children")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<HierarchyNode>? Children { get; set; }
    [JsonIgnore] public List<string>? Path { get; set; }
}

public sealed class PageRecord
{
    [JsonPropertyName("page_id")] public string PageId { get; set; } = "";
    [JsonPropertyName("title")] public string Title { get; set; } = "";
    [JsonPropertyName("text")] public string Text { get; set; } = "";
    [JsonIgnore] public List<string>? Path { get; set; }
    [JsonIgnore] public string? LastModifiedTime { get; set; }
}

public sealed class ActionItem
{
    [JsonPropertyName("page_id")] public string PageId { get; set; } = "";
    [JsonPropertyName("page_title")] public string PageTitle { get; set; } = "";
    [JsonPropertyName("source")] public string Source { get; set; } = "";
    [JsonPropertyName("text")] public string Text { get; set; } = "";
    [JsonPropertyName("owner")] public string? Owner { get; set; }
    [JsonPropertyName("due")] public string? Due { get; set; }
}

public sealed class InsightItem
{
    [JsonPropertyName("page_id")] public string PageId { get; set; } = "";
    [JsonPropertyName("page_title")] public string PageTitle { get; set; } = "";
    [JsonPropertyName("source")] public string Source { get; set; } = "";
    [JsonPropertyName("text")] public string Text { get; set; } = "";
}

public sealed class Insights
{
    [JsonPropertyName("action_items")] public List<ActionItem> ActionItems { get; set; } = new();
    [JsonPropertyName("decisions")] public List<InsightItem> Decisions { get; set; } = new();
    [JsonPropertyName("risks")] public List<InsightItem> Risks { get; set; } = new();
    [JsonPropertyName("questions")] public List<InsightItem> Questions { get; set; } = new();
    [JsonPropertyName("missing_owners")] public int MissingOwners { get; set; }
    [JsonPropertyName("missing_due_dates")] public int MissingDueDates { get; set; }
}

public sealed class PageRef
{
    [JsonPropertyName("page_id")] public string PageId { get; set; } = "";
    [JsonPropertyName("title")] public string Title { get; set; } = "";
}

public sealed class StalePageRef
{
    [JsonPropertyName("page_id")] public string PageId { get; set; } = "";
    [JsonPropertyName("title")] public string Title { get; set; } = "";
    [JsonPropertyName("last_modified")] public string? LastModified { get; set; }
}

public sealed class DuplicateCandidate
{
    [JsonPropertyName("left")] public string Left { get; set; } = "";
    [JsonPropertyName("right")] public string Right { get; set; } = "";
    [JsonPropertyName("titles")] public List<string> Titles { get; set; } = new();
    [JsonPropertyName("confidence")] public double Confidence { get; set; }
    [JsonPropertyName("reason")] public string Reason { get; set; } = "";
}

public sealed class ActionItemHealth
{
    [JsonPropertyName("total")] public int Total { get; set; }
    [JsonPropertyName("missing_owners")] public int MissingOwners { get; set; }
    [JsonPropertyName("missing_due_dates")] public int MissingDueDates { get; set; }
}

public sealed class HealthReport
{
    [JsonPropertyName("generated_at")] public string GeneratedAt { get; set; } = "";
    [JsonPropertyName("page_count")] public int PageCount { get; set; }
    [JsonPropertyName("organization_score")] public int OrganizationScore { get; set; }
    [JsonPropertyName("untitled_pages")] public List<PageRef> UntitledPages { get; set; } = new();
    [JsonPropertyName("empty_pages")] public List<PageRef> EmptyPages { get; set; } = new();
    [JsonPropertyName("stale_pages")] public List<StalePageRef> StalePages { get; set; } = new();
    [JsonPropertyName("duplicate_candidates")] public List<DuplicateCandidate> DuplicateCandidates { get; set; } = new();
    [JsonPropertyName("action_item_health")] public ActionItemHealth ActionItemHealthInfo { get; set; } = new();
    [JsonPropertyName("recommendations")] public List<string> Recommendations { get; set; } = new();
}

public sealed class DigestPage
{
    [JsonPropertyName("page_id")] public string PageId { get; set; } = "";
    [JsonPropertyName("title")] public string Title { get; set; } = "";
    [JsonPropertyName("path")] public List<string> Path { get; set; } = new();
    [JsonPropertyName("key_points")] public List<string> KeyPoints { get; set; } = new();
    [JsonPropertyName("word_count")] public int WordCount { get; set; }
}

public sealed class KnowledgeDigest
{
    [JsonPropertyName("mode")] public string Mode { get; set; } = "";
    [JsonPropertyName("page_count")] public int PageCount { get; set; }
    [JsonPropertyName("total_words")] public int TotalWords { get; set; }
    [JsonPropertyName("pages")] public List<DigestPage> Pages { get; set; } = new();
    [JsonPropertyName("insights")] public Insights InsightsInfo { get; set; } = new();
    [JsonPropertyName("synthesis_instruction")] public string SynthesisInstruction { get; set; } = "";
}

public static class KnowledgeOperator
{
    private const RegexOptions CI = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;

    private static readonly Regex Action = new(
        @"^(?:[-*•]\s*)?(?:\[[ xX]?\]\s*)?(?:action(?: item)?|todo|to-do|follow[- ]?up|next step|owner)\s*[:\-]\s*(.+)$", CI);
    private static readonly Regex Decision = new(
        @"^(?:[-*•]\s*)?(?:decision|decided|agreed|approved)\s*[:\-]\s*(.+)$", CI);
    private static readonly Regex Risk = new(
        @"^(?:[-*•]\s*)?(?:risk|blocker|issue|concern)\s*[:\-]\s*(.+)$", CI);
    private static readonly Regex Question = new(
        @"^(?:[-*•]\s*)?(?:question|open question)\s*[:\-]\s*(.+)$", CI);
    private static readonly Regex Owner = new(
        @"(?:^|\s)(?:owner|assignee)\s*[:=]\s*([^,;|]+?)(?=\s+(?:due|deadline)\s*[:=]|$)", CI);
    private static readonly Regex DueDate = new(
        @"\b(?:20\d{2}-\d{2}-\d{2}|(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2}(?:,\s*20\d{2})?)\b", CI);

    public static List<HierarchyNode> FlattenPages(IEnumerable<HierarchyNode> nodes, List<string>? path = null)
    {
        path ??= new List<string>();
        var pages = new List<HierarchyNode>();
        foreach (var node in nodes)
        {
            var nextPath = node.Kind == "Page"
                ? path
                : new List<string>(path) { string.IsNullOrEmpty(node.Name) ? node.Kind : node.Name };
            if (node.Kind == "Page")
            {
                pages.Add(new HierarchyNode
                {
                    Kind = node.Kind,
                    Id = node.Id,
                    Name = string.IsNullOrEmpty(node.Name) ? "Untitled" : node.Name,
                    CurrentlyViewed = node.CurrentlyViewed,
                    LastModifiedTime = node.LastModifiedTime,
                    DateTime = node.DateTime,
                    Children = null,
                    Path = path,
                });
            }
            if (node.Children is not null) pages.AddRange(FlattenPages(node.Children, nextPath));
        }
        return pages;
    }

    private static IEnumerable<string> Lines(string text) =>
        Regex.Split(text, @"\r?\n").Select(l => l.Trim()).Where(l => l.Length > 0);

    public static Insights ExtractInsights(IEnumerable<PageRecord> pages)
    {
        var result = new Insights();
        foreach (var page in pages)
        {
            foreach (var line in Lines(page.Text))
            {
                var action = Action.Match(line);
                if (action.Success)
                {
                    var owner = Owner.Match(line) is { Success: true } o ? o.Groups[1].Value.Trim() : null;
                    var due = DueDate.Match(line) is { Success: true } d ? d.Value : null;
                    result.ActionItems.Add(new ActionItem
                    {
                        PageId = page.PageId, PageTitle = page.Title, Source = line,
                        Text = action.Groups[1].Value.Trim(),
                        Owner = string.IsNullOrEmpty(owner) ? null : owner,
                        Due = string.IsNullOrEmpty(due) ? null : due,
                    });
                }
                else if (Decision.Match(line) is { Success: true } dec)
                {
                    result.Decisions.Add(new InsightItem { PageId = page.PageId, PageTitle = page.Title, Source = line, Text = dec.Groups[1].Value });
                }
                else if (Risk.Match(line) is { Success: true } risk)
                {
                    result.Risks.Add(new InsightItem { PageId = page.PageId, PageTitle = page.Title, Source = line, Text = risk.Groups[1].Value });
                }
                else if (Question.Match(line) is { Success: true } q)
                {
                    result.Questions.Add(new InsightItem { PageId = page.PageId, PageTitle = page.Title, Source = line, Text = q.Groups[1].Value });
                }
                else if (line.EndsWith('?'))
                {
                    result.Questions.Add(new InsightItem { PageId = page.PageId, PageTitle = page.Title, Source = line, Text = line });
                }
            }
        }
        result.MissingOwners = result.ActionItems.Count(item => item.Owner is null);
        result.MissingDueDates = result.ActionItems.Count(item => item.Due is null);
        return result;
    }

    private static HashSet<string> Tokens(string value) =>
        Regex.Replace(value.ToLowerInvariant(), @"[^a-z0-9\s]", " ")
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Where(word => word.Length > 2)
            .ToHashSet();

    public static double Similarity(string left, string right)
    {
        var a = Tokens(left);
        var b = Tokens(right);
        if (a.Count == 0 && b.Count == 0) return 1;
        var intersection = a.Count(word => b.Contains(word));
        var union = a.Union(b).Count();
        return union > 0 ? (double)intersection / union : 0;
    }

    public static HealthReport BuildHealthReport(IReadOnlyList<PageRecord> pages, DateTimeOffset now, int staleDays = 180)
    {
        var duplicates = new List<DuplicateCandidate>();
        for (var left = 0; left < pages.Count; left += 1)
        {
            for (var right = left + 1; right < pages.Count; right += 1)
            {
                var score = Similarity($"{pages[left].Title}\n{pages[left].Text}", $"{pages[right].Title}\n{pages[right].Text}");
                if (score >= 0.72)
                {
                    duplicates.Add(new DuplicateCandidate
                    {
                        Left = pages[left].PageId, Right = pages[right].PageId,
                        Titles = new List<string> { pages[left].Title, pages[right].Title },
                        Confidence = Math.Round(score, 3), Reason = "high token overlap",
                    });
                }
            }
        }
        var cutoff = now.ToUnixTimeMilliseconds() - (long)staleDays * 86_400_000L;
        var stale = pages.Where(page =>
            page.LastModifiedTime is not null
            && DateTimeOffset.TryParse(page.LastModifiedTime, null, System.Globalization.DateTimeStyles.AssumeUniversal, out var parsed)
            && parsed.ToUnixTimeMilliseconds() < cutoff).ToList();
        var empty = pages.Where(page => page.Text.Trim().Length == 0).ToList();
        var untitled = pages.Where(page =>
            page.Title.Trim().Length == 0 || Regex.IsMatch(page.Title.Trim(), "^untitled$", CI)).ToList();
        var insights = ExtractInsights(pages);
        var score2 = Math.Max(0,
            100 - untitled.Count * 4 - empty.Count * 3 - stale.Count * 2 - duplicates.Count * 5 - insights.MissingOwners * 2);
        var recommendations = new List<string>();
        if (duplicates.Count > 0) recommendations.Add($"Review {duplicates.Count} duplicate candidate pair(s).");
        if (untitled.Count > 0) recommendations.Add($"Rename {untitled.Count} untitled page(s).");
        if (stale.Count > 0) recommendations.Add($"Review {stale.Count} page(s) older than {staleDays} days.");
        if (insights.MissingOwners > 0) recommendations.Add($"Assign owners to {insights.MissingOwners} action item(s).");
        return new HealthReport
        {
            GeneratedAt = now.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'"),
            PageCount = pages.Count,
            OrganizationScore = score2,
            UntitledPages = untitled.Select(page => new PageRef { PageId = page.PageId, Title = page.Title }).ToList(),
            EmptyPages = empty.Select(page => new PageRef { PageId = page.PageId, Title = page.Title }).ToList(),
            StalePages = stale.Select(page => new StalePageRef { PageId = page.PageId, Title = page.Title, LastModified = page.LastModifiedTime }).ToList(),
            DuplicateCandidates = duplicates,
            ActionItemHealthInfo = new ActionItemHealth
            {
                Total = insights.ActionItems.Count,
                MissingOwners = insights.MissingOwners,
                MissingDueDates = insights.MissingDueDates,
            },
            Recommendations = recommendations,
        };
    }

    private static List<string> FirstSentences(string text, int count) =>
        Regex.Split(Regex.Replace(text, @"\s+", " "), @"(?<=[.!?])\s+")
            .Select(item => item.Trim())
            .Where(item => item.Length > 20)
            .Take(count)
            .ToList();

    public static KnowledgeDigest BuildKnowledgeDigest(IReadOnlyList<PageRecord> pages, string mode = "executive")
    {
        var insights = ExtractInsights(pages);
        var perPage = pages.Select(page => new DigestPage
        {
            PageId = page.PageId,
            Title = page.Title,
            Path = page.Path ?? new List<string>(),
            KeyPoints = FirstSentences(page.Text, mode == "detailed" ? 5 : 2),
            WordCount = page.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length,
        }).ToList();
        return new KnowledgeDigest
        {
            Mode = mode,
            PageCount = pages.Count,
            TotalWords = perPage.Sum(page => page.WordCount),
            Pages = perPage,
            InsightsInfo = insights,
            SynthesisInstruction = "Synthesize these source-grounded key points and structured insights. Cite page titles and do not invent facts absent from the records.",
        };
    }

    public static readonly IReadOnlyDictionary<string, Func<string, string, string>> Templates =
        new Dictionary<string, Func<string, string, string>>
        {
            ["meeting"] = (title, date) => $"# {title}\nDate: {date}\nAttendees:\n\n## Agenda\n\n## Notes\n\n## Decisions\n\n## Action Items\n- Action:  Owner:  Due:",
            ["project"] = (title, date) => $"# {title}\nUpdated: {date}\n\n## Objective\n\n## Status\n\n## Milestones\n\n## Risks and Blockers\n\n## Decisions\n\n## Next Steps",
            ["decision_log"] = (title, date) => $"# {title}\nUpdated: {date}\n\n## Decision\n\n## Context\n\n## Options Considered\n\n## Rationale\n\n## Owner\n\n## Review Date",
            ["weekly_review"] = (title, date) => $"# {title}\nWeek of: {date}\n\n## Wins\n\n## Decisions\n\n## Open Action Items\n\n## Risks\n\n## Priorities for Next Week",
        };

    public static string RenderTemplate(string template, string title, string? date = null)
    {
        if (!Templates.TryGetValue(template, out var factory))
            throw new ArgumentException($"Unknown template '{template}'. Available: {string.Join(", ", Templates.Keys)}");
        return factory(title, date ?? DateTime.UtcNow.ToString("yyyy-MM-dd"));
    }
}
