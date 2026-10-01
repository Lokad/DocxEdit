using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

// Shared patch-local identity: XML snapshot marks bind pre-existing objects
// and creation marks bind objects made by the patch. Image placement order comes
// from the single shared placement enumeration. Deleted objects keep historical
// identity in reports and must not resolve to survivors.

// One mark sweep shared by alias resolution, creation reporting, and
// preview binding. Story parts iterate in canonical order and the first
// match in document order wins. Callers keep their element filter and the
// meaning of the match; the helper only removes the duplicated loops.
internal static partial class DocxPatchEngine
{
    private static (StoryPartRef Story, XDocument Document, XElement Element)? FindMarkedStoryElement(
        OoxmlPackage package,
        XName markName,
        string mark,
        XName? elementName,
        CancellationToken cancellationToken)
    {
        foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(story.PartName);
            if (part is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, story.PartName, cancellationToken, out _);
            IEnumerable<XElement> candidates = elementName is null
                ? document.Descendants()
                : document.Descendants(elementName);
            XElement? match = candidates.FirstOrDefault(element =>
                string.Equals((string?)element.Attribute(markName), mark, StringComparison.Ordinal));
            if (match is not null)
            {
                return (story, document, match);
            }
        }

        return null;
    }
}
