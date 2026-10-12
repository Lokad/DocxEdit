using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

// Package-owned role discovery, rebuilt after relationship/content-type or
// membership changes. Story order and expected-root precedence stay explicit.
internal sealed class DocxPartRoleIndex
{
    public Dictionary<string, string> RelatedParts { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, XName> Roots { get; } = new(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyList<StoryPartRef> MainStory { get; }
    public IReadOnlyList<StoryPartRef> Stories { get; }
    public IReadOnlyList<string> WordParts { get; }

    public DocxPartRoleIndex(OoxmlPackage package, CancellationToken cancellationToken)
    {
        OoxmlRelationship[] relationships = package.GetRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
            .OrderBy(r => r.Id, StringComparer.Ordinal).ToArray();
        foreach (OoxmlRelationship relationship in relationships)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (package.GetPart(relationship.ResolvedTarget!) is not null)
                RelatedParts.TryAdd(relationship.Type, relationship.ResolvedTarget!);
        }
        Roots.Add(package.MainDocumentPartName, OoxmlNs.W + "document");
        foreach (var (type, root) in new (string, XName)[]
        {
            (OoxmlRelTypes.Header, OoxmlNs.W + "hdr"), (OoxmlRelTypes.Footer, OoxmlNs.W + "ftr"),
            (OoxmlRelTypes.Styles, OoxmlNs.W + "styles"), (OoxmlRelTypes.Numbering, OoxmlNs.W + "numbering"),
            (OoxmlRelTypes.Settings, OoxmlNs.W + "settings"), (OoxmlRelTypes.Comments, OoxmlNs.W + "comments"),
            (OoxmlRelTypes.CommentsExtended, OoxmlNs.W15 + "commentsEx"), (OoxmlRelTypes.CommentsIds, OoxmlNs.W16Cid + "commentsIds"),
            (OoxmlRelTypes.Footnotes, OoxmlNs.W + "footnotes"), (OoxmlRelTypes.Endnotes, OoxmlNs.W + "endnotes")
        })
        {
            if (type == OoxmlRelTypes.Header || type == OoxmlRelTypes.Footer)
            {
                foreach (OoxmlRelationship relationship in relationships.Where(r => r.Type == type)) Roots.TryAdd(relationship.ResolvedTarget!, root);
            }
            else if (RelatedParts.TryGetValue(type, out string? partName)) Roots.TryAdd(partName, root);
        }
        MainStory = [new("M", package.MainDocumentPartName, "main")];
        var stories = new List<StoryPartRef>(MainStory);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { package.MainDocumentPartName };
        AddStories(OoxmlRelTypes.Header, "H", "header");
        AddStories(OoxmlRelTypes.Footer, "F", "footer");
        Stories = stories;
        var wordParts = new List<string>();
        seen.Clear();
        foreach (StoryPartRef story in Stories) AddWordPart(story.PartName);
        foreach (string type in new[] { OoxmlRelTypes.Comments, OoxmlRelTypes.CommentsExtended, OoxmlRelTypes.CommentsIds, OoxmlRelTypes.Footnotes, OoxmlRelTypes.Endnotes })
            if (RelatedParts.TryGetValue(type, out string? name)) AddWordPart(name);
        foreach (OoxmlPart part in package.Parts.Values.Where(p => DocxPartRoles.IsConventionalWordXml(p.Name)).OrderBy(p => p.Name, StringComparer.Ordinal)) AddWordPart(part.Name);
        WordParts = wordParts;

        void AddStories(string type, string prefix, string label)
        {
            int index = 1;
            foreach (OoxmlRelationship relationship in relationships.Where(r => r.Type == type))
            {
                if (package.GetPart(relationship.ResolvedTarget!) is null || !seen.Add(relationship.ResolvedTarget!)) continue;
                stories.Add(new StoryPartRef($"{prefix}{index:000}", relationship.ResolvedTarget!, $"{label}[{index}]"));
                index++;
            }
        }
        void AddWordPart(string name)
        {
            if (package.GetPart(name) is not null && seen.Add(name)) wordParts.Add(name);
        }
    }
}
