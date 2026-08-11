// Port of test/knowledgeOperator.test.ts from the TypeScript source.

using OneNoteMcp;
using Xunit;

namespace OneNoteMcp.Tests;

public class KnowledgeOperatorTests
{
    private static List<PageRecord> Pages() => new()
    {
        new PageRecord
        {
            PageId = "p1",
            Title = "Client meeting",
            Text = "Decision: ship Friday.\nAction: Send contract Owner: Ana Due: 2026-08-15\nRisk: legal review is blocked",
            LastModifiedTime = "2025-01-01T00:00:00Z",
        },
        new PageRecord
        {
            PageId = "p2",
            Title = "Client meeting copy",
            Text = "Decision: ship Friday.\nAction: Send contract Owner: Ana Due: 2026-08-15\nRisk: legal review is blocked",
        },
        new PageRecord
        {
            PageId = "p3",
            Title = "Untitled",
            Text = "Action: confirm budget",
        },
    };

    [Fact]
    public void ExtractsStructuredDecisionsRisksActionOwnersAndDueDates()
    {
        var result = KnowledgeOperator.ExtractInsights(Pages());
        Assert.Equal(2, result.Decisions.Count);
        Assert.Equal(2, result.Risks.Count);
        Assert.Equal("Ana", result.ActionItems[0].Owner);
        Assert.Equal("2026-08-15", result.ActionItems[0].Due);
        Assert.Equal(1, result.MissingOwners);
    }

    [Fact]
    public void HealthReportExplainsDuplicatesStalenessUntitledPagesAndOwnerGaps()
    {
        var result = KnowledgeOperator.BuildHealthReport(
            Pages(), DateTimeOffset.Parse("2026-08-11T00:00:00Z"), 180);
        Assert.Single(result.DuplicateCandidates);
        Assert.Single(result.StalePages);
        Assert.Single(result.UntitledPages);
        Assert.True(result.Recommendations.Count >= 3);
    }

    [Fact]
    public void DigestStaysSourceGroundedAndExposesSynthesisInstructions()
    {
        var result = KnowledgeOperator.BuildKnowledgeDigest(Pages(), "detailed");
        Assert.Equal(3, result.PageCount);
        Assert.Matches("(?i)do not invent", result.SynthesisInstruction);
    }

    [Fact]
    public void SimilarityAndTemplatesAreDeterministic()
    {
        Assert.True(KnowledgeOperator.Similarity("alpha beta gamma", "alpha beta gamma delta") > 0.7);
        Assert.Matches("## Action Items", KnowledgeOperator.RenderTemplate("meeting", "Project Sync", "2026-08-11"));
        var ex = Assert.Throws<ArgumentException>(() => KnowledgeOperator.RenderTemplate("missing", "Nope"));
        Assert.Matches("Unknown template", ex.Message);
    }
}
