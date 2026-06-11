using System.IO.Compression;
using System.Xml.Linq;

namespace DocxEdit.Ooxml;

internal sealed class OoxmlPackage
{
    private readonly Dictionary<string, OoxmlPart> parts;

    private OoxmlPackage(Dictionary<string, OoxmlPart> parts, OoxmlContentTypes contentTypes, string? mainDocumentPartName)
    {
        this.parts = parts;
        ContentTypes = contentTypes;
        MainDocumentPartName = mainDocumentPartName;
        ContentTypesPart = parts["/[Content_Types].xml"];
    }

    public IReadOnlyDictionary<string, OoxmlPart> Parts => parts;
    public OoxmlPart ContentTypesPart { get; }
    public OoxmlContentTypes ContentTypes { get; }
    public string? MainDocumentPartName { get; }

    public static OoxmlPackage Load(
        Stream input,
        OoxmlPackageOptions options,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(options);
        if (!input.CanRead)
        {
            throw new ArgumentException("Input stream must be readable.", nameof(input));
        }

        cancellationToken.ThrowIfCancellationRequested();
        Stream archiveStream = input;
        MemoryStream? copy = null;
        if (!input.CanSeek)
        {
            copy = new MemoryStream();
            CopyTo(input, copy, options.MaxUncompressedBytes, cancellationToken);
            copy.Position = 0;
            archiveStream = copy;
        }

        try
        {
            using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read, leaveOpen: options.LeaveInputOpen || copy is not null);
            if (archive.Entries.Count > options.MaxZipEntries)
            {
                throw new InvalidDataException($"OOXML package has too many ZIP entries: {archive.Entries.Count}.");
            }

            ZipArchiveEntry contentTypesEntry = archive.GetEntry("[Content_Types].xml")
                ?? throw new InvalidDataException("OOXML package is missing [Content_Types].xml.");

            using Stream contentTypesStream = contentTypesEntry.Open();
            OoxmlContentTypes contentTypes = OoxmlContentTypes.Parse(contentTypesStream, cancellationToken);

            long totalBytes = 0;
            var parts = new Dictionary<string, OoxmlPart>(StringComparer.OrdinalIgnoreCase);

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
                    totalBytes += entry.Length;
                }

                if (totalBytes > options.MaxUncompressedBytes)
                {
                    throw new InvalidDataException("OOXML package exceeds the maximum supported uncompressed size.");
                }

                if (parts.ContainsKey(partName))
                {
                    throw new InvalidDataException($"OOXML package contains duplicate part '{partName}'.");
                }

                using Stream entryStream = entry.Open();
                using var memory = new MemoryStream(entry.Length > int.MaxValue ? 0 : (int)entry.Length);
                CopyTo(entryStream, memory, options.MaxSinglePartBytes, cancellationToken);

                string? contentType = partName == "/[Content_Types].xml"
                    ? OoxmlContentTypeNames.Xml
                    : contentTypes.GetContentType(partName);
                parts[partName] = new OoxmlPart(partName, entry.FullName, contentType, memory.ToArray());
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

            return new OoxmlPackage(parts, contentTypes, mainDocumentPartName);
        }
        finally
        {
            copy?.Dispose();
            if (!options.LeaveInputOpen && archiveStream != input)
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
        CancellationToken cancellationToken = default)
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

    public void Save(Stream output, CancellationToken cancellationToken = default)
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
    }

    internal void AddPart(string partName, string contentType, byte[] bytes)
    {
        string normalized = OoxmlPath.NormalizePartName(partName);
        if (parts.ContainsKey(normalized))
        {
            throw new InvalidDataException($"OOXML part '{normalized}' already exists.");
        }

        parts[normalized] = new OoxmlPart(normalized, normalized.TrimStart('/'), contentType, bytes);
        AddContentTypeOverride(normalized, contentType);
    }

    internal void AddRelationship(string sourcePartName, string relationshipId, string relationshipType, string target)
    {
        string relationshipPartName = OoxmlPath.GetRelationshipPartName(sourcePartName);
        XDocument document;
        if (parts.TryGetValue(relationshipPartName, out OoxmlPart? relationshipPart))
        {
            using Stream stream = relationshipPart.OpenRead();
            document = SafeXml.Load(stream);
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
        root.Add(new XElement(
            OoxmlNs.Rel + "Relationship",
            new XAttribute("Id", relationshipId),
            new XAttribute("Type", relationshipType),
            new XAttribute("Target", target)));

        using var output = new MemoryStream();
        document.Save(output, SaveOptions.DisableFormatting);
        ReplacePartBytes(relationshipPartName, output.ToArray());
    }

    private void AddContentTypeOverride(string partName, string contentType)
    {
        using Stream stream = ContentTypesPart.OpenRead();
        XDocument document = SafeXml.Load(stream);
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

    internal static IReadOnlyList<OoxmlRelationship> ParseRelationships(
        Stream stream,
        string sourcePartName,
        CancellationToken cancellationToken = default)
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
            string sourcePartName = GetSourcePartNameFromRelationshipPartName(relationshipPart.Name);
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

    private static string GetSourcePartNameFromRelationshipPartName(string relationshipPartName)
    {
        string normalized = OoxmlPath.NormalizePartName(relationshipPartName);
        if (normalized == "/_rels/.rels")
        {
            return "/";
        }

        const string relationshipMarker = "/_rels/";
        int markerIndex = normalized.LastIndexOf(relationshipMarker, StringComparison.Ordinal);
        if (markerIndex < 0 || !normalized.EndsWith(".rels", StringComparison.Ordinal))
        {
            throw new InvalidDataException($"Invalid relationship part name '{relationshipPartName}'.");
        }

        string directory = normalized[..markerIndex];
        string fileName = normalized[(markerIndex + relationshipMarker.Length)..^".rels".Length];
        return OoxmlPath.NormalizePartName($"{directory}/{fileName}");
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
