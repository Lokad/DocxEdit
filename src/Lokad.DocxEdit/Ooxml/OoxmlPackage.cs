using System.IO.Compression;
using System.Xml.Linq;

namespace Lokad.DocxEdit.Ooxml;

internal sealed class OoxmlPackage
{
    private readonly Dictionary<string, OoxmlPart> parts;
    private readonly HashSet<string> touchedPartNames = new(StringComparer.OrdinalIgnoreCase);

    private OoxmlPackage(Dictionary<string, OoxmlPart> parts, string mainDocumentPartName)
    {
        this.parts = parts;
        MainDocumentPartName = mainDocumentPartName;
    }

    public IReadOnlyDictionary<string, OoxmlPart> Parts => parts;
    public IReadOnlyCollection<string> TouchedPartNames => touchedPartNames;
    public OoxmlPart ContentTypesPart => parts["/[Content_Types].xml"];
    /// <summary>Main document part name. Never null: <see cref="Load"/> throws when the package has no main document.</summary>
    public string MainDocumentPartName { get; }

    // Ownership matrix: Load accepts input ownership at entry (after null/readable
    // validation). The input is disposed on every exit - success, load failure,
    // quota failure, or cancellation - if and only if LeaveInputOpen is false;
    // otherwise it is left open. For non-seekable input a seekable copy is made;
    // the archive wraps the copy (never the input) and the copy is always
    // disposed here. Every part is buffered to memory, so nothing borrows the
    // streams after Load returns.
    public static OoxmlPackage Load(
        Stream input,
        OoxmlPackageOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        if (!input.CanRead)
        {
            throw new ArgumentException("Input stream must be readable.", nameof(input));
        }

        cancellationToken.ThrowIfCancellationRequested();
        Stream archiveStream = input;
        MemoryStream? copy = null;

        try
        {
            if (!input.CanSeek)
            {
                copy = new MemoryStream();
                CopyTo(input, copy, options.MaxUncompressedBytes, cancellationToken);
                copy.Position = 0;
                archiveStream = copy;
            }

            using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read, leaveOpen: true);
            if (archive.Entries.Count > options.MaxZipEntries)
            {
                throw new InvalidDataException($"OOXML package has too many ZIP entries: {archive.Entries.Count}.");
            }

            // Pass 1: declared sizes, totals, and duplicates from entry metadata
            // alone, before any part content is read or parsed.
            long declaredTotalBytes = 0;
            var seenPartNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (entry.FullName.EndsWith('/'))
                {
                    continue;
                }

                string partName = OoxmlPath.NormalizeZipEntryName(entry.FullName);
                if (entry.Length > options.MaxSinglePartBytes)
                {
                    throw new InvalidDataException($"OOXML part '{partName}' exceeds the maximum supported size.");
                }

                checked
                {
                    declaredTotalBytes += entry.Length;
                }

                if (declaredTotalBytes > options.MaxUncompressedBytes)
                {
                    throw new InvalidDataException("OOXML package exceeds the maximum supported uncompressed size.");
                }

                if (!seenPartNames.Add(partName))
                {
                    throw new InvalidDataException($"OOXML package contains duplicate part '{partName}'.");
                }
            }

            ZipArchiveEntry contentTypesEntry = archive.GetEntry("[Content_Types].xml")
                ?? throw new InvalidDataException("OOXML package is missing [Content_Types].xml.");

            OoxmlContentTypes contentTypes;
            using (Stream contentTypesStream = contentTypesEntry.Open())
            using (var contentTypesBuffer = new MemoryStream())
            {
                CopyTo(contentTypesStream, contentTypesBuffer, options.MaxSinglePartBytes, cancellationToken);
                contentTypesBuffer.Position = 0;
                contentTypes = OoxmlContentTypes.Parse(contentTypesBuffer, cancellationToken);
            }

            // Pass 2: buffer every part through bounded copies while enforcing
            // the actual (not just declared) total size.
            long actualTotalBytes = 0;
            var parts = new Dictionary<string, OoxmlPart>(StringComparer.OrdinalIgnoreCase);

            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (entry.FullName.EndsWith('/'))
                {
                    continue;
                }

                string partName = OoxmlPath.NormalizeZipEntryName(entry.FullName);
                using Stream entryStream = entry.Open();
                using var memory = new MemoryStream(entry.Length > int.MaxValue ? 0 : (int)entry.Length);
                CopyTo(entryStream, memory, options.MaxSinglePartBytes, cancellationToken);
                byte[] partBytes = memory.ToArray();

                checked
                {
                    actualTotalBytes += partBytes.Length;
                }

                if (actualTotalBytes > options.MaxUncompressedBytes)
                {
                    throw new InvalidDataException("OOXML package exceeds the maximum supported uncompressed size.");
                }

                string? contentType = partName == "/[Content_Types].xml"
                    ? OoxmlContentTypeNames.Xml
                    : contentTypes.GetContentType(partName);
                parts[partName] = new OoxmlPart(partName, entry.FullName, contentType, partBytes);
            }

            if (!parts.ContainsKey("/[Content_Types].xml"))
            {
                throw new InvalidDataException("OOXML package is missing [Content_Types].xml.");
            }

            string mainDocumentPartName = FindMainDocumentPartName(parts, cancellationToken);
            ValidateInternalRelationshipTargets(parts, cancellationToken);
            if (!options.AllowMacroEnabledDocuments &&
                parts[mainDocumentPartName].ContentType == OoxmlContentTypeNames.MacroEnabledMainDocument)
            {
                throw new InvalidDataException("Macro-enabled Word documents are not allowed.");
            }

            ValidateMainDocumentNamespace(parts[mainDocumentPartName], cancellationToken);

            string mainDocumentRelationshipsPartName = OoxmlPath.GetRelationshipPartName(mainDocumentPartName);
            if (!parts.TryGetValue(mainDocumentRelationshipsPartName, out OoxmlPart? mainDocumentRelationshipsPart))
            {
                throw new InvalidDataException($"OOXML package is missing {mainDocumentRelationshipsPartName}.");
            }

            using (Stream relationshipStream = mainDocumentRelationshipsPart.OpenRead())
            {
                _ = ParseRelationships(relationshipStream, mainDocumentPartName, cancellationToken);
            }

            return new OoxmlPackage(parts, mainDocumentPartName);
        }
        finally
        {
            copy?.Dispose();
            if (!options.LeaveInputOpen)
            {
                input.Dispose();
            }
        }
    }

    public OoxmlPart? GetPart(string partName)
    {
        return parts.TryGetValue(OoxmlPath.NormalizePartName(partName), out OoxmlPart? part) ? part : null;
    }

    public IReadOnlyList<OoxmlRelationship> GetRelationships(
        string sourcePartName,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string relationshipPartName = OoxmlPath.GetRelationshipPartName(sourcePartName);
        OoxmlPart? relationshipPart = GetPart(relationshipPartName);
        if (relationshipPart is null)
        {
            return [];
        }

        using Stream stream = relationshipPart.OpenRead();
        return ParseRelationships(stream, sourcePartName, cancellationToken);
    }

    /// <summary>Enumerates package-internal relationships with a resolved part-name target.</summary>
    public IReadOnlyList<ResolvedOoxmlRelationship> GetResolvedRelationships(
        string sourcePartName,
        CancellationToken cancellationToken)
    {
        var resolved = new List<ResolvedOoxmlRelationship>();
        foreach (OoxmlRelationship relationship in GetRelationships(sourcePartName, cancellationToken))
        {
            if (relationship.IsExternal || relationship.ResolvedTarget is null)
            {
                continue;
            }

            resolved.Add(new ResolvedOoxmlRelationship(relationship.Id, relationship.Type, relationship.ResolvedTarget));
        }

        return resolved;
    }

    public void Save(Stream output, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(output);
        if (!output.CanWrite)
        {
            throw new ArgumentException("Output stream must be writable.", nameof(output));
        }

        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);
        foreach (OoxmlPart part in parts.Values.OrderBy(part => part.OriginalEntryName, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string entryName = part.OriginalEntryName.Length == 0
                ? part.Name.TrimStart('/')
                : part.OriginalEntryName;
            ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using Stream stream = entry.Open();
            stream.Write(part.Bytes, 0, part.Bytes.Length);
        }
    }

    internal void ReplacePartBytes(string partName, byte[] bytes)
    {
        string normalized = OoxmlPath.NormalizePartName(partName);
        if (!parts.TryGetValue(normalized, out OoxmlPart? part))
        {
            throw new InvalidDataException($"OOXML part '{normalized}' does not exist.");
        }

        parts[normalized] = part with { Bytes = bytes };
        touchedPartNames.Add(normalized);
    }

    // D01: snapshot bookkeeping must not mark otherwise-untouched parts as edited.
    // Capture annotates story parts transiently; untouched parts are restored
    // byte-identical afterwards so output preserves original bytes.
    internal void UnmarkPartTouched(string partName)
    {
        touchedPartNames.Remove(OoxmlPath.NormalizePartName(partName));
    }

    internal void AddPart(string partName, string contentType, byte[] bytes, CancellationToken cancellationToken)
    {
        string normalized = OoxmlPath.NormalizePartName(partName);
        if (parts.ContainsKey(normalized))
        {
            throw new InvalidDataException($"OOXML part '{normalized}' already exists.");
        }

        parts[normalized] = new OoxmlPart(normalized, normalized.TrimStart('/'), contentType, bytes);
        touchedPartNames.Add(normalized);
        AddContentTypeOverride(normalized, contentType, cancellationToken);
    }

    internal void AddRelationship(string sourcePartName, string relationshipId, string relationshipType, string target, string? targetMode, CancellationToken cancellationToken)
    {
        string relationshipPartName = OoxmlPath.GetRelationshipPartName(sourcePartName);
        XDocument document;
        if (parts.TryGetValue(relationshipPartName, out OoxmlPart? relationshipPart))
        {
            using Stream stream = relationshipPart.OpenRead();
            document = SafeXml.Load(stream, cancellationToken);
        }
        else
        {
            document = new XDocument(new XElement(OoxmlNs.Rel + "Relationships"));
            parts[relationshipPartName] = new OoxmlPart(
                relationshipPartName,
                relationshipPartName.TrimStart('/'),
                OoxmlContentTypeNames.Relationships,
                []);
        }

        XElement root = document.Root
            ?? throw new InvalidDataException($"Relationship part '{relationshipPartName}' has no XML root.");
        var relationship = new XElement(
            OoxmlNs.Rel + "Relationship",
            new XAttribute("Id", relationshipId),
            new XAttribute("Type", relationshipType),
            new XAttribute("Target", target));
        if (!string.IsNullOrWhiteSpace(targetMode))
        {
            relationship.SetAttributeValue("TargetMode", targetMode);
        }

        root.Add(relationship);

        using var output = new MemoryStream();
        document.Save(output, SaveOptions.DisableFormatting);
        ReplacePartBytes(relationshipPartName, output.ToArray());
    }

    internal void RemoveRelationship(string sourcePartName, string relationshipId, CancellationToken cancellationToken)
    {
        string relationshipPartName = OoxmlPath.GetRelationshipPartName(sourcePartName);
        if (!parts.TryGetValue(relationshipPartName, out OoxmlPart? relationshipPart))
        {
            return;
        }

        using Stream stream = relationshipPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        foreach (XElement relationship in document.Root?.Elements(OoxmlNs.Rel + "Relationship")
            .Where(element => string.Equals((string?)element.Attribute("Id"), relationshipId, StringComparison.Ordinal))
            .ToArray() ?? [])
        {
            relationship.Remove();
        }

        using var output = new MemoryStream();
        document.Save(output, SaveOptions.DisableFormatting);
        ReplacePartBytes(relationshipPartName, output.ToArray());
    }

    internal void RemovePart(string partName, CancellationToken cancellationToken)
    {
        string normalized = OoxmlPath.NormalizePartName(partName);
        if (!parts.Remove(normalized))
        {
            return;
        }

        touchedPartNames.Remove(normalized);
        RemoveContentTypeOverride(normalized, cancellationToken);
    }

    private void AddContentTypeOverride(string partName, string contentType, CancellationToken cancellationToken)
    {
        using Stream stream = ContentTypesPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        XElement root = document.Root
            ?? throw new InvalidDataException("Content types part has no XML root.");
        bool exists = root
            .Elements(OoxmlNs.Ct + "Override")
            .Any(element => string.Equals((string?)element.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase));
        if (!exists)
        {
            root.Add(new XElement(
                OoxmlNs.Ct + "Override",
                new XAttribute("PartName", partName),
                new XAttribute("ContentType", contentType)));
        }

        using var output = new MemoryStream();
        document.Save(output, SaveOptions.DisableFormatting);
        ReplacePartBytes("/[Content_Types].xml", output.ToArray());
    }

    private void RemoveContentTypeOverride(string partName, CancellationToken cancellationToken)
    {
        using Stream stream = ContentTypesPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        XElement root = document.Root
            ?? throw new InvalidDataException("Content types part has no XML root.");
        foreach (XElement element in root
            .Elements(OoxmlNs.Ct + "Override")
            .Where(element => string.Equals((string?)element.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase))
            .ToArray())
        {
            element.Remove();
        }

        using var output = new MemoryStream();
        document.Save(output, SaveOptions.DisableFormatting);
        ReplacePartBytes("/[Content_Types].xml", output.ToArray());
    }

    internal static IReadOnlyList<OoxmlRelationship> ParseRelationships(
        Stream stream,
        string sourcePartName,
        CancellationToken cancellationToken)
    {
        XDocument document = SafeXml.Load(stream, cancellationToken);
        var relationships = new List<OoxmlRelationship>();
        var ids = new HashSet<string>(StringComparer.Ordinal);

        foreach (XElement element in document.Root?.Elements(OoxmlNs.Rel + "Relationship") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string id = RequiredAttribute(element, "Id");
            if (!ids.Add(id))
            {
                throw new InvalidDataException($"Duplicate relationship Id '{id}' in '{OoxmlPath.GetRelationshipPartName(sourcePartName)}'.");
            }

            string type = RequiredAttribute(element, "Type");
            string target = RequiredAttribute(element, "Target");
            string? targetMode = (string?)element.Attribute("TargetMode");
            string? resolvedTarget = targetMode?.Equals("External", StringComparison.OrdinalIgnoreCase) == true
                ? null
                : OoxmlPath.ResolveRelationshipTarget(sourcePartName, target);

            relationships.Add(new OoxmlRelationship(id, type, target, targetMode, resolvedTarget));
        }

        return relationships;
    }

    private static string FindMainDocumentPartName(
        Dictionary<string, OoxmlPart> parts,
        CancellationToken cancellationToken)
    {
        if (!parts.TryGetValue("/_rels/.rels", out OoxmlPart? packageRelationshipsPart))
        {
            throw new InvalidDataException("OOXML package is missing /_rels/.rels.");
        }

        using Stream stream = packageRelationshipsPart.OpenRead();
        OoxmlRelationship? officeDocumentRelationship = ParseRelationships(stream, "/", cancellationToken)
            .FirstOrDefault(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.OfficeDocument);
        if (officeDocumentRelationship?.ResolvedTarget is null)
        {
            throw new InvalidDataException("OOXML package is missing the office document relationship.");
        }

        if (!parts.ContainsKey(officeDocumentRelationship.ResolvedTarget))
        {
            throw new InvalidDataException($"Main document part '{officeDocumentRelationship.ResolvedTarget}' does not exist.");
        }

        return officeDocumentRelationship.ResolvedTarget;
    }

    private static void ValidateInternalRelationshipTargets(
        Dictionary<string, OoxmlPart> parts,
        CancellationToken cancellationToken)
    {
        foreach (OoxmlPart relationshipPart in parts.Values.Where(part => part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string sourcePartName = OoxmlPath.GetSourcePartNameFromRelationshipPartName(relationshipPart.Name);
            using Stream stream = relationshipPart.OpenRead();
            foreach (OoxmlRelationship relationship in ParseRelationships(stream, sourcePartName, cancellationToken))
            {
                if (!relationship.IsExternal &&
                    relationship.ResolvedTarget is not null &&
                    !parts.ContainsKey(relationship.ResolvedTarget))
                {
                    throw new InvalidDataException($"Relationship '{relationship.Id}' in '{relationshipPart.Name}' targets missing part '{relationship.ResolvedTarget}'.");
                }
            }
        }
    }

    private static void ValidateMainDocumentNamespace(OoxmlPart mainDocumentPart, CancellationToken cancellationToken)
    {
        using Stream stream = mainDocumentPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        if (document.Root?.Name != OoxmlNs.W + "document")
        {
            throw new InvalidDataException($"Main document part '{mainDocumentPart.Name}' is not a WordprocessingML document.");
        }
    }

    private static void CopyTo(
        Stream source,
        Stream destination,
        long maxBytes,
        CancellationToken cancellationToken)
    {
        byte[] buffer = new byte[81920];
        long total = 0;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0)
            {
                return;
            }

            total += read;
            if (total > maxBytes)
            {
                throw new InvalidDataException("OOXML stream exceeds the maximum supported size.");
            }

            destination.Write(buffer, 0, read);
        }
    }

    private static string RequiredAttribute(XElement element, string name)
    {
        return (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required Relationship attribute '{name}'.");
    }
}
