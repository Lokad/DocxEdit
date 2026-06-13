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

