using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private sealed record CommentReplyAnchor(string PartName, XDocument Document, XElement? Start, XElement? End, XElement Reference);

    private static CommentReplyAnchor? FindCommentReplyAnchor(OoxmlPackage package, string parentId, CancellationToken cancellationToken)
    {
        CommentReplyAnchor? found = null;
        foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, story.PartName, cancellationToken, out _);
            XElement[] references = document.Descendants(OoxmlNs.W + "commentReference")
                .Where(e => (string?)e.Attribute(OoxmlNs.W + "id") == parentId).ToArray();
            if (references.Length == 0) continue;
            if (found is not null || references.Length != 1 || references[0].Parent?.Name != OoxmlNs.W + "r") return null;
            XElement[] starts = document.Descendants(OoxmlNs.W + "commentRangeStart")
                .Where(e => (string?)e.Attribute(OoxmlNs.W + "id") == parentId).ToArray();
            XElement[] ends = document.Descendants(OoxmlNs.W + "commentRangeEnd")
                .Where(e => (string?)e.Attribute(OoxmlNs.W + "id") == parentId).ToArray();
            if (starts.Length > 1 || starts.Length != ends.Length) return null;
            found = new CommentReplyAnchor(story.PartName, document, starts.FirstOrDefault(), ends.FirstOrDefault(), references[0]);
        }
        return found;
    }

    private static void AddReplyAnchor(CommentReplyAnchor anchor, string replyId)
    {
        anchor.Start?.AddAfterSelf(new XElement(OoxmlNs.W + "commentRangeStart", new XAttribute(OoxmlNs.W + "id", replyId)));
        anchor.End?.AddAfterSelf(new XElement(OoxmlNs.W + "commentRangeEnd", new XAttribute(OoxmlNs.W + "id", replyId)));
        anchor.Reference.Parent!.AddAfterSelf(new XElement(OoxmlNs.W + "r",
            new XElement(OoxmlNs.W + "commentReference", new XAttribute(OoxmlNs.W + "id", replyId))));
    }

    private sealed record CommentAnchorTarget(string PartName, XDocument Document, XElement Container, XElement[] Paragraphs);

    private static CommentAnchorTarget? ResolveCommentAnchorTarget(
        OoxmlPackage package, DocxPatchOperation operation, string target,
        CancellationToken cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (DocxTargetId.TryParse(target, out DocxTargetId id) && id.Kind is DocxTargetKind.Cell or DocxTargetKind.MergeGroup)
        {
            CellTarget? cell = ResolveCellTarget(package, target, cancellationToken);
            if (cell is null) return null;
            if (IsVerticalMergeContinuation(cell.Cell))
            {
                diagnostics = [Diagnostic(DocxSeverity.Error, "E4301", $"Unsupported merged-cell target '{target}'. Target the vertical-merge root cell instead.", operation, target)];
                return null;
            }
            XElement[] paragraphs = cell.Cell.Elements(OoxmlNs.W + "p").ToArray();
            if (paragraphs.Length == 0 || cell.Cell.Elements().Any(e => e.Name != OoxmlNs.W + "p" && e.Name != OoxmlNs.W + "tcPr"))
            {
                diagnostics = [Diagnostic(DocxSeverity.Error, "E4317", $"Comment target '{target}' must contain direct paragraphs only; nested tables and block wrappers are unsupported.", operation, target)];
                return null;
            }
            return new CommentAnchorTarget(cell.PartName, cell.Document, cell.Cell, paragraphs);
        }
        ParagraphTarget? paragraph = ResolveParagraphTarget(package, operation, target, cancellationToken, out diagnostics);
        return paragraph is null ? null : new CommentAnchorTarget(paragraph.PartName, paragraph.Document, paragraph.Paragraph, [paragraph.Paragraph]);
    }
}
