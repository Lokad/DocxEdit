namespace DocxEdit;

public interface IDocxAssetProvider
{
    bool TryOpen(
        string reference,
        out Stream stream,
        out string? contentTypeHint,
        out string? fileNameHint);
}

