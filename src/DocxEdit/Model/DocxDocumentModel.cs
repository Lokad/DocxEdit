namespace DocxEdit.Model;

internal sealed record DocxDocumentModel(
    IReadOnlyList<DocxParagraphInfo> Paragraphs,
    IReadOnlyList<DocxTableInfo> Tables,
    IReadOnlyList<DocxImageInfo> Images,
    IReadOnlyList<DocxSectionInfo> Sections,
    IReadOnlyList<DocxBookmarkInfo> Bookmarks,
    IReadOnlyList<DocxContentControlInfo> ContentControls,
    IReadOnlyList<DocxFieldInfo> Fields)
{
    public static DocxDocumentModel Empty { get; } = new([], [], [], [], [], [], []);
}
