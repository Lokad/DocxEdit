namespace Lokad.DocxEdit;

// One immutable snapshot per distinct decoded reference for one command. All
// external I/O finishes asynchronously before the synchronous engine consumes it.
internal sealed class BufferedDocxAssets : IDocxAssetProvider
{
    private sealed record Entry(byte[] Bytes, string? ContentType, string? FileName, Exception? ReadError);
    private readonly Dictionary<string, Entry?> entries = new(StringComparer.Ordinal);

    public static async Task<BufferedDocxAssets> LoadAsync(
        DocxPatch patch, IDocxAsyncAssetProvider provider, long maximum, CancellationToken cancellationToken)
    {
        var buffered = new BufferedDocxAssets();
        // The parser has validated field names against the operation registry
        // and decoded quotes/heredocs. Do not parse asset syntax a second time.
        foreach (DocxPatchOperation operation in patch.Operations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!operation.Fields.TryGetValue("asset", out string? reference) ||
                reference.Length == 0 || buffered.entries.ContainsKey(reference)) continue;

            try
            {
                DocxAsset? asset = await provider.OpenAsync(reference, cancellationToken).ConfigureAwait(false);
                if (asset is null)
                {
                    buffered.entries[reference] = null;
                    continue;
                }
                await using Stream source = asset.Content;
                using var contents = new MemoryStream();
                await DocxCommandExecution.CopyToBoundedAsync(source, contents, maximum, cancellationToken).ConfigureAwait(false);
                buffered.entries[reference] = new Entry(contents.ToArray(), asset.ContentTypeHint, asset.FileNameHint, null);
            }
            catch (Exception exception) when (exception is IOException or InvalidDataException)
            {
                // Preserve the engine's diagnostic and operation ordering: an
                // unreachable asset error must not replace an earlier guard error.
                buffered.entries[reference] = new Entry([], null, null, exception);
            }
        }
        return buffered;
    }

    public bool TryOpen(string reference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)
    {
        stream = Stream.Null;
        contentTypeHint = null;
        fileNameHint = null;
        if (!entries.TryGetValue(reference, out Entry? entry) || entry is null) return false;
        stream = entry.ReadError is { } error ? new FailedReadStream(error) : new MemoryStream(entry.Bytes, writable: false);
        contentTypeHint = entry.ContentType;
        fileNameHint = entry.FileName;
        return true;
    }

    private sealed class FailedReadStream(Exception error) : MemoryStream
    {
        public override int Read(byte[] buffer, int offset, int count) => throw error;
        public override int Read(Span<byte> buffer) => throw error;
    }
}
