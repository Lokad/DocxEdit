namespace Lokad.DocxEdit.Ooxml;

internal static class OoxmlPath
{
    public static string NormalizePartName(string partName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(partName);

        if (partName.Contains('\\', StringComparison.Ordinal))
        {
            throw new InvalidDataException($"Unsafe OOXML part path '{partName}'.");
        }

        string normalized = partName;
        if (!normalized.StartsWith('/'))
        {
            normalized = "/" + normalized;
        }

        string[] segments = normalized.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0)
        {
            throw new InvalidDataException("OOXML part path cannot refer to the package root.");
        }

        foreach (string segment in segments)
        {
            if (segment is "." or ".." || segment.Contains(':', StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsafe OOXML part path '{partName}'.");
            }
        }

        return "/" + string.Join('/', segments);
    }

    public static string NormalizeZipEntryName(string entryName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(entryName);
        if (entryName.StartsWith('/') || entryName.StartsWith('\\') || entryName.Contains('\\', StringComparison.Ordinal))
        {
            throw new InvalidDataException($"Unsafe OOXML ZIP entry path '{entryName}'.");
        }

        return NormalizePartName(entryName);
    }

    public static string ResolveRelationshipTarget(string sourcePartName, string target)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(target);

        if (target.Contains(':', StringComparison.Ordinal) || target.Contains('\\', StringComparison.Ordinal))
        {
            throw new InvalidDataException($"External or unsafe relationship target '{target}' is not a package part.");
        }

        if (target.StartsWith('/'))
        {
            return NormalizePartName(target);
        }

        string normalizedSource = sourcePartName == "/" ? "/" : NormalizePartName(sourcePartName);
        string baseDirectory = normalizedSource == "/"
            ? "/"
            : normalizedSource[..(normalizedSource.LastIndexOf('/') + 1)];

        var segments = new List<string>();
        if (baseDirectory.Length > 1)
        {
            segments.AddRange(baseDirectory.Split('/', StringSplitOptions.RemoveEmptyEntries));
        }

        foreach (string segment in target.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".")
            {
                continue;
            }

            if (segment == "..")
            {
                if (segments.Count == 0)
                {
                    throw new InvalidDataException($"Relationship target '{target}' escapes the OOXML package root.");
                }

                segments.RemoveAt(segments.Count - 1);
                continue;
            }

            if (segment.Contains(':', StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsafe relationship target '{target}'.");
            }

            segments.Add(segment);
        }

        if (segments.Count == 0)
        {
            throw new InvalidDataException($"Relationship target '{target}' does not name a package part.");
        }

        return "/" + string.Join('/', segments);
    }

    public static string GetRelationshipPartName(string sourcePartName)
    {
        if (sourcePartName == "/")
        {
            return "/_rels/.rels";
        }

        string normalized = NormalizePartName(sourcePartName);
        int slash = normalized.LastIndexOf('/');
        string directory = normalized[..(slash + 1)];
        string fileName = normalized[(slash + 1)..];
        return $"{directory}_rels/{fileName}.rels";
    }

    /// <summary>Maps a relationship part name back to its source part.</summary>
    public static string GetSourcePartNameFromRelationshipPartName(string relationshipPartName)
    {
        string normalized = NormalizePartName(relationshipPartName);
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
        return NormalizePartName($"{directory}/{fileName}");
    }
}

