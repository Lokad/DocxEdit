namespace Lokad.DocxEdit;

/// <summary>Provides patch-referenced binary assets (for example images for <c>replace-image</c>) to the patch engine.</summary>
/// <remarks>The CLI resolves references against the file system; library consumers supply their own source (archives, stores, tests).</remarks>
public interface IDocxAssetProvider
{
    /// <summary>Tries to open <paramref name="reference"/> for reading.</summary>
    /// <remarks>Return <c>false</c> when the reference cannot be resolved; the engine turns that into a patch diagnostic, never an exception. Do not throw for a missing asset.</remarks>
    /// <param name="reference">Caller-visible asset reference as written in the patch (for example a file path).</param>
    /// <param name="stream">Open readable stream on success; value is ignored on <c>false</c>.</param>
    /// <param name="contentTypeHint">Optional content-type override; null lets the engine infer it.</param>
    /// <param name="fileNameHint">Optional file name used for part naming and diagnostics.</param>
    /// <returns><c>true</c> when the asset was opened; <c>false</c> otherwise.</returns>
    bool TryOpen(
        string reference,
        out Stream stream,
        out string? contentTypeHint,
        out string? fileNameHint);
}
