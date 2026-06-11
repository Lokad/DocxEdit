namespace DocxEdit.Model;

internal sealed record DocxDocumentModel(
    IReadOnlyList<DocxParagraphInfo> Paragraphs,
    IReadOnlyList<DocxTableInfo> Tables,
    IReadOnlyList<DocxImageInfo> Images,
    IReadOnlyList<DocxSectionInfo> Sections)
{
    public static DocxDocumentModel Empty { get; } = new([], [], [], []);
}
