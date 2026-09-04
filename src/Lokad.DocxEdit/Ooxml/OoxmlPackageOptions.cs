namespace Lokad.DocxEdit.Ooxml;

/// <summary>Package-load quotas and policy. All components are mandatory; typical values come from <see cref="Lokad.DocxEdit.DocxPackageLimits"/>.</summary>
internal sealed record OoxmlPackageOptions(
    bool LeaveInputOpen,
    int MaxZipEntries,
    long MaxUncompressedBytes,
    long MaxSinglePartBytes,
    bool AllowMacroEnabledDocuments);
