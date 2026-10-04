namespace Lokad.DocxEdit.Model;

internal sealed record DocxDocumentModel(
    IReadOnlyList<DocxParagraphInfo> Paragraphs,
    IReadOnlyList<DocxTableInfo> Tables,
    IReadOnlyList<DocxImageInfo> Images,
    IReadOnlyList<DocxSectionInfo> Sections,
    IReadOnlyList<DocxBookmarkInfo> Bookmarks,
    IReadOnlyList<DocxContentControlInfo> ContentControls,
    IReadOnlyList<DocxFieldInfo> Fields,
    IReadOnlyList<DocxHyperlinkInfo> Hyperlinks)
{
    public IReadOnlyList<DocxEquationInfo> Equations { get; init; } = [];
    public static DocxDocumentModel Empty { get; } = new([], [], [], [], [], [], [], []);
}
