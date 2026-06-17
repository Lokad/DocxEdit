namespace Lokad.DocxEdit.Ooxml;

internal sealed record OoxmlPart(string Name, string OriginalEntryName, string? ContentType, byte[] Bytes)
{
    public Stream OpenRead()
    {
        return new MemoryStream(Bytes, writable: false);
    }
}

