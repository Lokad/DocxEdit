namespace DocxEdit.Model;

internal sealed record DocxDocumentModel(
    IReadOnlyList<DocxParagraphInfo> Paragraphs,
    IReadOnlyList<DocxTableInfo> Tables,
    IReadOnlyList<DocxImageInfo> Images)
{
    public static DocxDocumentModel Empty { get; } = new([], [], []);
}

