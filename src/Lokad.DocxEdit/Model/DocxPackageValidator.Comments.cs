using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxPackageValidator
{
    private static void ValidateCommentsExtended(
        OoxmlPackage package,
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        HashSet<string> commentParaIds = ReadCommentParaIds(package, cancellationToken);
        var extensionParaIds = document
            .Descendants(OoxmlNs.W15 + "commentEx")
            .Select(element => (ParaId: (string?)element.Attribute(OoxmlNs.W15 + "paraId"), Element: element))
            .ToArray();

        foreach ((string? paraId, XElement _) in extensionParaIds.Where(item => string.IsNullOrWhiteSpace(item.ParaId)))
        {
            diagnostics.Add(Error("E9108", "commentsExtended commentEx is missing w15:paraId.", partName));
        }

        foreach (IGrouping<string, string> group in extensionParaIds
            .Select(item => item.ParaId)
            .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
            .OfType<string>()
            .GroupBy(paraId => paraId, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            diagnostics.Add(Error("E9108", $"Duplicate commentsExtended paraId '{group.Key}' appears {group.Count()} times.", partName));
        }

        foreach (string paraId in extensionParaIds
            .Select(item => item.ParaId)
            .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal))
        {
            if (!commentParaIds.Contains(paraId))
            {
                diagnostics.Add(Error("E9108", $"commentsExtended paraId '{paraId}' has no matching comment paragraph.", partName));
            }
        }

        foreach ((string? paraId, string? parentParaId) in document
            .Descendants(OoxmlNs.W15 + "commentEx")
            .Select(element => (
                ParaId: (string?)element.Attribute(OoxmlNs.W15 + "paraId"),
                ParentParaId: (string?)element.Attribute(OoxmlNs.W15 + "paraIdParent")))
            .Where(item => !string.IsNullOrWhiteSpace(item.ParentParaId)))
        {
            if (string.Equals(paraId, parentParaId, StringComparison.Ordinal))
            {
                diagnostics.Add(Error("E9108", $"commentsExtended paraId '{paraId}' cannot parent itself.", partName));
                continue;
            }

            if (parentParaId is null || !commentParaIds.Contains(parentParaId))
            {
                diagnostics.Add(Error("E9108", $"commentsExtended parent paraId '{parentParaId}' has no matching comment paragraph.", partName));
            }
        }
    }

    private static void ValidateCommentsIds(
        OoxmlPackage package,
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        HashSet<string> commentParaIds = ReadCommentParaIds(package, cancellationToken);
        var ids = document
            .Descendants(OoxmlNs.W16Cid + "commentId")
            .Select(element => (
                ParaId: (string?)element.Attribute(OoxmlNs.W16Cid + "paraId"),
                DurableId: (string?)element.Attribute(OoxmlNs.W16Cid + "durableId")))
            .ToArray();

        foreach ((string? paraId, _) in ids.Where(item => string.IsNullOrWhiteSpace(item.ParaId)))
        {
            diagnostics.Add(Error("E9122", "commentsIds commentId is missing w16cid:paraId.", partName));
        }

        foreach ((string? paraId, string? durableId) in ids.Where(item => string.IsNullOrWhiteSpace(item.DurableId)))
        {
            string target = string.IsNullOrWhiteSpace(paraId) ? "with no paraId" : $"for paraId '{paraId}'";
            diagnostics.Add(Error("E9122", $"commentsIds commentId {target} is missing w16cid:durableId.", partName));
        }

        foreach (IGrouping<string, string> group in ids
            .Select(item => item.ParaId)
            .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
            .OfType<string>()
            .GroupBy(paraId => paraId, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            diagnostics.Add(Error("E9122", $"Duplicate commentsIds paraId '{group.Key}' appears {group.Count()} times.", partName));
        }

        foreach (IGrouping<string, string> group in ids
            .Select(item => item.DurableId)
            .Where(durableId => !string.IsNullOrWhiteSpace(durableId))
            .OfType<string>()
            .GroupBy(durableId => durableId, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            diagnostics.Add(Error("E9122", $"Duplicate commentsIds durableId '{group.Key}' appears {group.Count()} times.", partName));
        }

        foreach (string paraId in ids
            .Select(item => item.ParaId)
            .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal))
        {
            if (!commentParaIds.Contains(paraId))
            {
                diagnostics.Add(Error("E9122", $"commentsIds paraId '{paraId}' has no matching comment paragraph.", partName));
            }
        }
    }

    private static void ValidateSettings(
        XDocument document,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        foreach (XElement updateFields in document.Descendants(OoxmlNs.W + "updateFields"))
        {
            string? value = (string?)updateFields.Attribute(OoxmlNs.W + "val");
            if (value is null || value is "0" or "1" or "true" or "false" or "on" or "off")
            {
                continue;
            }

            diagnostics.Add(Error("E9118", $"Settings w:updateFields has invalid OnOff value '{value}'.", partName));
        }
    }

    private static HashSet<string> ReadCommentParaIds(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var paraIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            OoxmlPart? commentsPart = package.GetPart(partName);
            if (commentsPart is null)
            {
                continue;
            }

            using Stream stream = commentsPart.OpenRead();
            XDocument comments = SafeXml.Load(stream, cancellationToken);
            foreach (string paraId in comments
                .Descendants(OoxmlNs.W + "comment")
                .Select(comment => (string?)comment.Descendants(OoxmlNs.W + "p").LastOrDefault()?.Attribute(OoxmlNs.W14 + "paraId"))
                .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
                .OfType<string>())
            {
                paraIds.Add(paraId);
            }
        }

        return paraIds;
    }

    private static void ValidateCommentConsistency(
        OoxmlPackage package,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        var bodyIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XElement comment in document.Descendants(OoxmlNs.W + "comment"))
            {
                string? id = (string?)comment.Attribute(OoxmlNs.W + "id");
                if (string.IsNullOrWhiteSpace(id))
                {
                    diagnostics.Add(Error("E9111", "Comment body is missing w:id.", partName));
                    continue;
                }

                if (!bodyIds.TryGetValue(id, out List<string>? partNames))
                {
                    partNames = [];
                    bodyIds[id] = partNames;
                }

                partNames.Add(partName);
            }
        }

        foreach (KeyValuePair<string, List<string>> item in bodyIds
            .Where(item => item.Value.Count > 1)
            .OrderBy(item => item.Key, StringComparer.Ordinal))
        {
            string partNames = string.Join(", ", item.Value.Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal));
            diagnostics.Add(Error("E9111", $"Duplicate comment body id '{item.Key}' appears {item.Value.Count} times in comments parts: {partNames}.", item.Value[0]));
        }

        var rangeIds = new HashSet<string>(StringComparer.Ordinal);
        var referenceIds = new HashSet<string>(StringComparer.Ordinal);
        string? commentsPartName = DocxPartRoles.FindCommentsPartName(package, cancellationToken);
        string? commentsExtendedPartName = DocxPartRoles.FindCommentsExtendedPartName(package, cancellationToken);
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) &&
                part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !part.Name.Contains("/_rels/", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(part.Name, commentsPartName, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(part.Name, commentsExtendedPartName, StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XElement marker in document
                .Descendants()
                .Where(element => element.Name == OoxmlNs.W + "commentRangeStart" ||
                    element.Name == OoxmlNs.W + "commentRangeEnd" ||
                    element.Name == OoxmlNs.W + "commentReference"))
            {
                string? id = (string?)marker.Attribute(OoxmlNs.W + "id");
                if (string.IsNullOrWhiteSpace(id))
                {
                    diagnostics.Add(Error("E9111", $"{marker.Name.LocalName} is missing w:id.", part.Name));
                    continue;
                }

                if (marker.Name == OoxmlNs.W + "commentReference")
                {
                    referenceIds.Add(id);
                }
                else
                {
                    rangeIds.Add(id);
                }
            }
        }

        foreach (string id in rangeIds.Concat(referenceIds).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal))
        {
            if (!bodyIds.ContainsKey(id))
            {
                diagnostics.Add(Error("E9111", $"Comment markup id '{id}' has no matching comment body.", null));
            }
        }

        foreach (string id in rangeIds.Except(referenceIds, StringComparer.Ordinal).Order(StringComparer.Ordinal))
        {
            diagnostics.Add(Error("E9111", $"Comment range id '{id}' has no matching commentReference.", null));
        }
    }
}
