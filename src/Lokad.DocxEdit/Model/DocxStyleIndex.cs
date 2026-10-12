using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

// Package-owned style metadata. No XML tree or caller stream is retained;
// replacing the source bytes or relocating its relationship rebuilds the index.
internal sealed class DocxStyleIndex
{
    private readonly Dictionary<string, Definition> definitions = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int?> levels = new(StringComparer.Ordinal);
    private sealed record Definition(string? BasedOn, bool HasOutline, int? Level);

    public string? PartName { get; }
    public byte[]? SourceBytes { get; }
    public IReadOnlyList<DocxStyleInfo> Styles { get; }

    private DocxStyleIndex(string? partName, OoxmlPart? part, XElement? root, CancellationToken cancellationToken)
    {
        PartName = partName;
        SourceBytes = part?.Bytes;
        Styles = DocxStyleScanner.ReadStyles(root, cancellationToken);
        foreach (XElement style in root?.Elements(OoxmlNs.W + "style") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? id = (string?)style.Attribute(OoxmlNs.W + "styleId");
            if (id is null) continue;
            bool hasOutline = DocxHeadingLevels.TryReadOutlineLevel(style.Element(OoxmlNs.W + "pPr"), out int? level);
            // Heading lookup historically uses the first raw definition. The
            // scanner still reports duplicate supported styles at its boundary.
            definitions.TryAdd(id, new Definition((string?)style.Element(OoxmlNs.W + "basedOn")?.Attribute(OoxmlNs.W + "val"), hasOutline, level));
        }
    }

    public static DocxStyleIndex Load(string? partName, OoxmlPart? part, CancellationToken cancellationToken)
    {
        XElement? root = null;
        if (part is not null)
        {
            using Stream stream = part.OpenRead();
            root = SafeXml.Load(stream, cancellationToken).Root;
        }
        return new DocxStyleIndex(partName, part, root, cancellationToken);
    }

    public int? GetHeadingLevel(string styleId, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!definitions.ContainsKey(styleId))
            return DocxHeadingLevels.TryReadBuiltInLevel(styleId, out int builtInLevel) ? builtInLevel : null;
        if (levels.TryGetValue(styleId, out int? cached)) return cached;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        int? result = null;
        string? current = styleId;
        while (!string.IsNullOrEmpty(current) && visited.Add(current))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!definitions.TryGetValue(current, out Definition? definition)) break;
            if (levels.TryGetValue(current, out result)) break;
            if (definition.HasOutline) { result = definition.Level; break; }
            current = definition.BasedOn;
        }
        foreach (string id in visited)
        {
            // A missing built-in style only supplies a latent level when it is
            // requested directly, never when a defined style inherits from it.
            if (definitions.ContainsKey(id)) levels[id] = result;
        }
        return result;
    }
}
