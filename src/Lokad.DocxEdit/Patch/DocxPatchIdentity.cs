using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

// C02 and C03 shared identity contract. Work in progress. Placement
// enumeration is the single owner for discovery, snapshot, mutation,
// and reporting. Public ordinal IDs are views, not identity.
// A patch local handle names one object for the duration of a patch.
// Input handles bind pre patch objects. Created handles bind objects
// made by the patch. Deleted objects are historical and must not
// resolve to survivors. Do not infer identity twice from neighbors.

internal sealed record PatchObjectHandle(
    DocxTargetKind Kind,
    char Story,
    int StoryPart,
    int Ordinal,
    bool IsCreated,
    Guid InstanceId);

internal sealed record OperationTargetResolution(
    PatchObjectHandle Handle,
    XElement Element,
    string? SnapshotId);

// C04: one mark sweep shared by alias resolution, creation reporting, and
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
