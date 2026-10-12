using System.Text;
using System.Text.Json;
using Lokad.DocxEdit.Ooxml;

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

    internal static Task CopyToBoundedAsync(Stream source, MemoryStream output, long maximum, CancellationToken cancellationToken)
    {
        return BoundedStreams.CopyToAsync(source, output, maximum, "Command input exceeds the byte quota.", cancellationToken);
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

    private async Task FlushTextAsync(StringWriter buffered, TextWriter destination)
    {
        foreach (ReadOnlyMemory<char> chunk in buffered.GetStringBuilder().GetChunks())
        {
            _cancellationToken.ThrowIfCancellationRequested();
            await destination.WriteAsync(chunk, _cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task PublishJsonAsync<T>(string path, T value, JsonSerializerOptions options)
    {
        using var bytes = new MemoryStream();
        await JsonSerializer.SerializeAsync(bytes, value, options, _cancellationToken).ConfigureAwait(false);
        using var readable = new MemoryStream(bytes.GetBuffer(), 0, checked((int)bytes.Length), writable: false);
        _cancellationToken.ThrowIfCancellationRequested();
        await _host.PublishFileAsync(path, readable, _cancellationToken).ConfigureAwait(false);
    }
}
