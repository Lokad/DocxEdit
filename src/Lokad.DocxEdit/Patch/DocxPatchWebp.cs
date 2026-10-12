using System.Buffers.Binary;
using System.Text;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    // Bounded container/header validation, without decoding pixels or rewriting
    // any chunk. See https://developers.google.com/speed/webp/docs/riff_container.
    private static bool TryReadWebp(byte[] bytes, out int width, out int height, out string? reason)
    {
        width = height = 0;
        reason = null;
        if (bytes.Length < 20 || !IsWebpSignature(bytes))
            return WebpFailure("missing RIFF/WEBP header or image chunk", out reason);
        if (BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(4, 4)) != bytes.Length - 8 || (bytes.Length & 1) != 0)
            return WebpFailure("RIFF size does not match the file", out reason);

        bool extended = false, image = false, lossless = false, alpha = false;
        bool icc = false, exif = false, xmp = false;
        int canvasWidth = 0, canvasHeight = 0;
        byte flags = 0;
        int position = 12;
        while (position < bytes.Length)
        {
            if (bytes.Length - position < 8)
                return WebpFailure("truncated chunk header", out reason);
            uint size = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(position + 4, 4));
            long end = position + 8L + size;
            long next = end + (size & 1);
            if (next > bytes.Length)
                return WebpFailure("truncated chunk payload or padding", out reason);
            if ((size & 1) != 0 && bytes[(int)end] != 0)
                return WebpFailure("nonzero RIFF padding", out reason);
            string kind = Encoding.ASCII.GetString(bytes, position, 4);
            ReadOnlySpan<byte> payload = bytes.AsSpan(position + 8, (int)size);
            switch (kind)
            {
                case "VP8X":
                    if (position != 12 || extended || payload.Length != 10)
                        return WebpFailure("VP8X must be the first chunk with a ten-byte header", out reason);
                    extended = true;
                    flags = payload[0];
                    if ((flags & 2) != 0)
                        return WebpFailure("animated WebP is unsupported; supply a still image", out reason);
                    if ((flags & 0xC1) != 0 || payload[1] != 0 || payload[2] != 0 || payload[3] != 0)
                        return WebpFailure("unsupported VP8X reserved fields", out reason);
                    canvasWidth = ReadWebpUInt24(payload[4..]) + 1;
                    canvasHeight = ReadWebpUInt24(payload[7..]) + 1;
                    if ((long)canvasWidth * canvasHeight > uint.MaxValue)
                        return WebpFailure("WebP canvas exceeds the format limit", out reason);
                    break;
                case "ANIM":
                case "ANMF":
                    return WebpFailure("animated WebP is unsupported; supply a still image", out reason);
                case "ALPH":
                    if (!extended || (flags & 0x10) == 0 || alpha || image || payload.Length < 2)
                        return WebpFailure("invalid or misplaced alpha chunk", out reason);
                    if ((payload[0] & 0xC0) != 0 || (payload[0] & 3) > 1 || ((payload[0] >> 4) & 3) > 1)
                        return WebpFailure("unsupported alpha encoding", out reason);
                    if ((payload[0] & 3) == 0 && payload.Length != 1L + (long)canvasWidth * canvasHeight)
                        return WebpFailure("raw alpha size does not match the canvas", out reason);
                    alpha = true;
                    break;
                case "VP8 ":
                    if (image || payload.Length <= 10 || !payload.Slice(3, 3).SequenceEqual(new byte[] { 0x9D, 0x01, 0x2A }))
                        return WebpFailure("missing or duplicate VP8 key-frame header", out reason);
                    int tag = ReadWebpUInt24(payload);
                    if ((tag & 1) != 0 || ((tag >> 1) & 7) > 3 || (tag & 0x10) == 0 ||
                        (tag >> 5) == 0 || 10L + (tag >> 5) > payload.Length)
                        return WebpFailure("invalid or truncated VP8 frame partition", out reason);
                    width = BinaryPrimitives.ReadUInt16LittleEndian(payload[6..]) & 0x3FFF;
                    height = BinaryPrimitives.ReadUInt16LittleEndian(payload[8..]) & 0x3FFF;
                    if (width == 0 || height == 0)
                        return WebpFailure("VP8 dimensions must be positive", out reason);
                    image = true;
                    break;
                case "VP8L":
                    if (image || alpha || payload.Length <= 5 || payload[0] != 0x2F)
                        return WebpFailure("missing or duplicate VP8L header or incompatible alpha chunk", out reason);
                    uint header = BinaryPrimitives.ReadUInt32LittleEndian(payload[1..]);
                    if ((header >> 29) != 0)
                        return WebpFailure("unsupported VP8L version", out reason);
                    width = (int)(header & 0x3FFF) + 1;
                    height = (int)((header >> 14) & 0x3FFF) + 1;
                    image = lossless = true;
                    break;
                case "ICCP":
                    if (!extended || icc || alpha || image || (flags & 0x20) == 0 || payload.IsEmpty)
                        return WebpFailure("invalid or misplaced ICC profile", out reason);
                    icc = true;
                    break;
                case "EXIF":
                    if (!extended || exif || (flags & 8) == 0 || payload.IsEmpty)
                        return WebpFailure("invalid Exif metadata chunk", out reason);
                    exif = true;
                    break;
                case "XMP ":
                    if (!extended || xmp || (flags & 4) == 0 || payload.IsEmpty)
                        return WebpFailure("invalid XMP metadata chunk", out reason);
                    xmp = true;
                    break;
                // Unknown chunks are retained byte for byte, as required by RIFF readers.
            }
            position = (int)next;
        }
        if (!image) return WebpFailure("missing still-image bitstream", out reason);
        if (extended && (width != canvasWidth || height != canvasHeight))
            return WebpFailure("still-image dimensions do not match the canvas", out reason);
        if (extended && (((flags & 0x20) != 0) != icc || ((flags & 8) != 0) != exif || ((flags & 4) != 0) != xmp ||
            ((flags & 0x10) != 0 && !lossless && !alpha)))
            return WebpFailure("VP8X features do not match the image chunks", out reason);
        return true;
    }

    private static bool IsWebpSignature(byte[] bytes) => bytes.Length >= 12 &&
        bytes.AsSpan(0, 4).SequenceEqual("RIFF"u8) && bytes.AsSpan(8, 4).SequenceEqual("WEBP"u8);

    private static int ReadWebpUInt24(ReadOnlySpan<byte> bytes) => bytes[0] | (bytes[1] << 8) | (bytes[2] << 16);

    private static bool WebpFailure(string message, out string? reason)
    {
        reason = message;
        return false;
    }
}
