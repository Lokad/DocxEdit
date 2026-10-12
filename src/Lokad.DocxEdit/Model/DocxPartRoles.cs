using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

// Central owner of package part-role discovery (PLAN B14). Every role resolves
// through package relationships; fixed paths are only ever used when CREATING
// parts, never to define existing ones. Consequences, all deliberate:
// - Relocated packages (live parts under unusual paths) behave like standard ones.
// - A part shared by several relationships of one kind is discovered once under
//   the first relationship index; duplicate work and duplicate IDs are gone.
// - Parts without a relationship role (orphans) receive no role-based checks.
// - Unsupported story kinds stay out of the read model: footnotes, endnotes,
//   and comment bodies are inventoried by changes and validated, but expose no
//   paragraph/table targets.
internal sealed record StoryPartRef(string Prefix, string PartName, string StoryLabel);

internal static class DocxPartRoles
{
    public static string? FindStylesPartName(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return FindRelatedPartName(package, OoxmlRelTypes.Styles, cancellationToken);
    }

    public static string? FindNumberingPartName(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return FindRelatedPartName(package, OoxmlRelTypes.Numbering, cancellationToken);
    }

    public static string? FindCommentsPartName(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return FindRelatedPartName(package, OoxmlRelTypes.Comments, cancellationToken);
    }

    public static string? FindCommentsExtendedPartName(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return FindRelatedPartName(package, OoxmlRelTypes.CommentsExtended, cancellationToken);
    }

    public static string? FindCommentsIdsPartName(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return FindRelatedPartName(package, OoxmlRelTypes.CommentsIds, cancellationToken);
    }

    public static string? FindSettingsPartName(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return FindRelatedPartName(package, OoxmlRelTypes.Settings, cancellationToken);
    }

    public static string? FindFootnotesPartName(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return FindRelatedPartName(package, OoxmlRelTypes.Footnotes, cancellationToken);
    }

    public static string? FindEndnotesPartName(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return FindRelatedPartName(package, OoxmlRelTypes.Endnotes, cancellationToken);
    }

    private static string? FindRelatedPartName(OoxmlPackage package, string relationshipType, CancellationToken cancellationToken)
    {
        return package.GetPartRoleIndex(cancellationToken).RelatedParts.GetValueOrDefault(relationshipType);
    }

    // Ordered story parts: main first, then distinct headers and footers in
    // relationship-Id order. Shared or dangling relationships consume no index.
    public static IReadOnlyList<StoryPartRef> GetOrderedStories(
        OoxmlPackage package,
        bool includeHeadersFooters,
        CancellationToken cancellationToken)
    {
        DocxPartRoleIndex index = package.GetPartRoleIndex(cancellationToken);
        return includeHeadersFooters ? index.Stories : index.MainStory;
    }

    public static IReadOnlyDictionary<string, string> GetStoryPrefixes(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        return GetOrderedStories(package, includeHeadersFooters: true, cancellationToken)
            .ToDictionary(story => story.PartName, story => story.Prefix, StringComparer.OrdinalIgnoreCase);
    }

    // Expected XML root by role; null when the part has no known role.
    public static XName? GetExpectedRoot(OoxmlPackage package, string partName, CancellationToken cancellationToken)
    {
        return package.GetPartRoleIndex(cancellationToken).Roots.GetValueOrDefault(partName);
    }

    // Distinct existing parts carrying word-processing markup for sweeps
    // (field refresh, revision IDs, comment IDs, post-edit range checks):
    // stories plus comment/footnote/endnote roles, plus conventional word XML
    // parts outside any role so orphans keep their legacy coverage.
    public static IReadOnlyList<string> GetWordProcessingParts(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return package.GetPartRoleIndex(cancellationToken).WordParts;
    }

    public static bool IsConventionalWordXml(string partName)
    {
        return partName.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) &&
            partName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
            !partName.Contains("/_rels/", StringComparison.OrdinalIgnoreCase);
    }

    // Parts validated as word-processing stories: conventional word XML plus
    // any role part (so relocated stories keep full structural validation).
    public static bool IsValidatedWordPart(OoxmlPackage package, string partName, CancellationToken cancellationToken)
    {
        if (IsConventionalWordXml(partName))
        {
            return true;
        }

        return GetWordProcessingParts(package, cancellationToken)
            .Contains(partName, StringComparer.OrdinalIgnoreCase);
    }
}
