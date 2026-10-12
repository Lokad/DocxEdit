using System.Buffers;

namespace Lokad.DocxEdit.Ooxml;

// Scratch buffers belong only to the copy operation. Package/result bytes
// remain independently owned, and quota checks use bytes actually read.
internal static class BoundedStreams
{
    private const int BufferSize = 81920;

    public static void CopyTo(Stream source, Stream destination, long maximum, string error, CancellationToken cancellationToken)
    {
        byte[] buffer = ArrayPool<byte>.Shared.Rent(BufferSize);
        try
        {
            long total = 0;
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(buffer, 0, BufferSize);
                if (read == 0) return;
                if (read > maximum - total) throw new InvalidDataException(error);
                total += read;
                destination.Write(buffer, 0, read);
            }
        }
        finally { ArrayPool<byte>.Shared.Return(buffer, clearArray: true); }
    }

    public static async Task CopyToAsync(Stream source, MemoryStream destination, long maximum, string error, CancellationToken cancellationToken)
    {
        byte[] buffer = ArrayPool<byte>.Shared.Rent(BufferSize);
        try
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int read = await source.ReadAsync(buffer.AsMemory(0, BufferSize), cancellationToken).ConfigureAwait(false);
                if (read == 0) return;
                if (read > maximum - destination.Length) throw new InvalidDataException(error);
                destination.Write(buffer, 0, read);
            }
        }
        finally { ArrayPool<byte>.Shared.Return(buffer, clearArray: true); }
    }
}
