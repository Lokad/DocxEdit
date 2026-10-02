using System.Text;

namespace Lokad.DocxEdit;

internal sealed partial class DocxCommandExecution
{
    private async Task<MemoryStream> OpenInputFileAsync(string path)
    {
        await using Stream source = await _host.OpenReadAsync(path, _cancellationToken).ConfigureAwait(false);
        var buffered = new MemoryStream();
        try
        {
            await CopyToBoundedAsync(source, buffered, _options.Quotas.MaxUncompressedBytes, _cancellationToken).ConfigureAwait(false);
            buffered.Position = 0;
            return buffered;
        }
        catch
        {
            buffered.Dispose();
            throw;
        }
    }

    internal static async Task CopyToBoundedAsync(Stream source, MemoryStream output, long maximum, CancellationToken cancellationToken)
    {
        var buffer = new byte[81920];
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = await source.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            if (output.Length + read > maximum)
                throw new InvalidDataException("Command input exceeds the byte quota.");
            output.Write(buffer, 0, read);
        }
    }

    private async Task<StringReader> OpenPatchFileAsync(string path)
    {
        using TextReader reader = await _host.OpenTextAsync(path, _cancellationToken).ConfigureAwait(false);
        // Retain at most one character beyond the parser quota so the engine
        // reports the existing E2014 diagnostic without consuming unbounded input.
        int maximum = new DocxEditOptions().MaxPatchChars;
        var text = new StringBuilder();
        var buffer = new char[8192];
        while (text.Length <= maximum)
        {
            int read = await reader.ReadAsync(buffer.AsMemory(0, Math.Min(buffer.Length, maximum + 1 - text.Length)), _cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            text.Append(buffer, 0, read);
        }
        return new StringReader(text.ToString());
    }

    private async Task PublishTextAsync(string path, string text)
    {
        using var bytes = new MemoryStream(Encoding.UTF8.GetBytes(text), writable: false);
        _cancellationToken.ThrowIfCancellationRequested();
        await _host.PublishFileAsync(path, bytes, _cancellationToken).ConfigureAwait(false);
    }
}
