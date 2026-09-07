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
        foreach (OoxmlRelationship relationship in package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => !relationship.IsExternal &&
                string.Equals(relationship.Type, relationshipType, StringComparison.Ordinal) &&
                relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            if (package.GetPart(relationship.ResolvedTarget!) is not null)
            {
                return relationship.ResolvedTarget;
            }
        }

        return null;
    }

    // Ordered story parts: main first, then distinct headers and footers in
    // relationship-Id order. Shared or dangling relationships consume no index.
    public static IReadOnlyList<StoryPartRef> GetOrderedStories(
        OoxmlPackage package,
        bool includeHeadersFooters,
        CancellationToken cancellationToken)
    {
        var stories = new List<StoryPartRef> { new("M", package.MainDocumentPartName, "main") };
        if (!includeHeadersFooters)
        {
            return stories;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { package.MainDocumentPartName };
        AddStoryParts(OoxmlRelTypes.Header, "H", "header");
        AddStoryParts(OoxmlRelTypes.Footer, "F", "footer");
        return stories;

        void AddStoryParts(string relationshipType, string prefixLetter, string storyLabel)
        {
            int index = 1;
            foreach (ResolvedOoxmlRelationship relationship in package
                .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
                .Where(candidate => candidate.Type == relationshipType)
                .OrderBy(candidate => candidate.Id, StringComparer.Ordinal))
            {
                if (package.GetPart(relationship.ResolvedTarget) is null || !seen.Add(relationship.ResolvedTarget))
                {
                    continue;
                }

                stories.Add(new StoryPartRef($"{prefixLetter}{index:000}", relationship.ResolvedTarget, $"{storyLabel}[{index}]"));
                index++;
            }
        }
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
        if (string.Equals(partName, package.MainDocumentPartName, StringComparison.OrdinalIgnoreCase))
        {
            return OoxmlNs.W + "document";
        }

        if (IsStoryKind(package, partName, OoxmlRelTypes.Header, cancellationToken))
        {
            return OoxmlNs.W + "hdr";
        }

        if (IsStoryKind(package, partName, OoxmlRelTypes.Footer, cancellationToken))
        {
            return OoxmlNs.W + "ftr";
        }

        if (string.Equals(partName, FindStylesPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            return OoxmlNs.W + "styles";
        }

        if (string.Equals(partName, FindNumberingPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            return OoxmlNs.W + "numbering";
        }

        if (string.Equals(partName, FindSettingsPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            return OoxmlNs.W + "settings";
        }

        if (string.Equals(partName, FindCommentsPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            return OoxmlNs.W + "comments";
        }

        if (string.Equals(partName, FindCommentsExtendedPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            return OoxmlNs.W15 + "commentsEx";
        }

        if (string.Equals(partName, FindCommentsIdsPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            return OoxmlNs.W16Cid + "commentsIds";
        }

        if (string.Equals(partName, FindFootnotesPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            return OoxmlNs.W + "footnotes";
        }

        if (string.Equals(partName, FindEndnotesPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            return OoxmlNs.W + "endnotes";
        }

        return null;
    }

    private static bool IsStoryKind(OoxmlPackage package, string partName, string relationshipType, CancellationToken cancellationToken)
    {
        return package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == relationshipType)
            .Any(relationship => string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase));
    }

    // Distinct existing parts carrying word-processing markup for sweeps
    // (field refresh, revision IDs, comment IDs, post-edit range checks):
    // stories plus comment/footnote/endnote roles, plus conventional word XML
    // parts outside any role so orphans keep their legacy coverage.
    public static IReadOnlyList<string> GetWordProcessingParts(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var parts = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        void Add(string? partName)
        {
            if (partName is not null && package.GetPart(partName) is not null && seen.Add(partName))
            {
                parts.Add(partName);
            }
        }

        Add(package.MainDocumentPartName);
        foreach (StoryPartRef story in GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            Add(story.PartName);
        }

        Add(FindCommentsPartName(package, cancellationToken));
        Add(FindCommentsExtendedPartName(package, cancellationToken));
        Add(FindCommentsIdsPartName(package, cancellationToken));
        Add(FindFootnotesPartName(package, cancellationToken));
        Add(FindEndnotesPartName(package, cancellationToken));
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => IsConventionalWordXml(part.Name))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            Add(part.Name);
        }

        return parts;
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
