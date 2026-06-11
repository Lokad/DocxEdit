namespace DocxEdit.Ooxml;

internal sealed record OoxmlPackageOptions(
    bool LeaveInputOpen = true,
    int MaxZipEntries = 10_000,
    long MaxUncompressedBytes = 512L * 1024L * 1024L,
    long MaxSinglePartBytes = 128L * 1024L * 1024L,
    bool AllowMacroEnabledDocuments = false);
