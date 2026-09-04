namespace Lokad.DocxEdit.Ooxml;

internal sealed record OoxmlRelationship(
    string Id,
    string Type,
    string Target,
    string? TargetMode,
    string? ResolvedTarget)
{
    public bool IsExternal => TargetMode?.Equals("External", StringComparison.OrdinalIgnoreCase) == true;
}

/// <summary>
/// A package-internal relationship whose target resolved to a part name.
/// External relationships never resolve, so resolved-only enumeration also
/// implies package-internal. Prefer this over null-forgiving access to
/// <see cref="OoxmlRelationship.ResolvedTarget"/>.
/// </summary>
internal sealed record ResolvedOoxmlRelationship(
    string Id,
    string Type,
    string ResolvedTarget);
