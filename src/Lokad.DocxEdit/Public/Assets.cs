namespace Lokad.DocxEdit;

/// <summary>Asynchronously opens patch-referenced assets for hosted command execution.</summary>
public interface IDocxAsyncAssetProvider
{
    /// <summary>Opens an asset, or returns null when it cannot be resolved. The caller owns and disposes the returned content stream. Cancellation must propagate.</summary>
    ValueTask<DocxAsset?> OpenAsync(string reference, CancellationToken cancellationToken);
}

/// <summary>An opened asset and optional media hints. Ownership of Content transfers to the caller of the asset provider.</summary>
/// <param name="Content">Owned readable stream; asynchronous-only reads are supported by hosted commands.</param>
/// <param name="ContentTypeHint">Optional media type; actual PNG/JPEG bytes remain authoritative.</param>
/// <param name="FileNameHint">Optional filename used for media validation and diagnostics.</param>
public sealed record DocxAsset(Stream Content, string? ContentTypeHint, string? FileNameHint);

/// <summary>Provides patch-referenced binary assets (for example images for <c>replace-image</c>) to the patch engine.</summary>
/// <remarks>The CLI resolves references against the file system; library consumers supply their own source (archives, stores, tests). The patch reader passed to check/apply stays caller-owned and is never disposed by the library.</remarks>
public interface IDocxAssetProvider
{
    /// <summary>Tries to open <paramref name="reference"/> for reading.</summary>
    /// <remarks>Return <c>false</c> when the reference cannot be resolved; the engine turns that into a patch diagnostic, never an exception. Do not throw for a missing asset.</remarks>
    /// <param name="reference">Caller-visible asset reference as written in the patch (for example a file path).</param>
    /// <param name="stream">Open readable stream on success; value is ignored on <c>false</c>. Streams opened this way are owned by the engine, which disposes them after reading on success, failure, and cancellation.</param>
    /// <param name="contentTypeHint">Optional content-type override; null lets the engine infer it.</param>
    /// <param name="fileNameHint">Optional file name used for part naming and diagnostics.</param>
    /// <returns><c>true</c> when the asset was opened; <c>false</c> otherwise.</returns>
    bool TryOpen(
        string reference,
        out Stream stream,
        out string? contentTypeHint,
        out string? fileNameHint);
}
