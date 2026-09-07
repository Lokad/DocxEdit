namespace Lokad.DocxEdit.Tests;

public static class DocxTargetIdTests
{
    // kindValue follows DocxTargetKind declaration order.
    [Theory]
    [InlineData("M.P0001", 'M', 0, DocxTargetKind.Paragraph, 1, 0, 0)]
    [InlineData("M.T0001", 'M', 0, DocxTargetKind.Table, 1, 0, 0)]
    [InlineData("M.T0001.R02", 'M', 0, DocxTargetKind.Row, 1, 2, 0)]
    [InlineData("M.T0001.R02.C03", 'M', 0, DocxTargetKind.Cell, 1, 2, 3)]
    [InlineData("M.T0001.MG0001", 'M', 0, DocxTargetKind.MergeGroup, 1, 1, 0)]
    [InlineData("M.I0001", 'M', 0, DocxTargetKind.Image, 1, 0, 0)]
    [InlineData("M.L0001", 'M', 0, DocxTargetKind.Hyperlink, 1, 0, 0)]
    [InlineData("M.F0001", 'M', 0, DocxTargetKind.Field, 1, 0, 0)]
    [InlineData("M.B0001", 'M', 0, DocxTargetKind.Bookmark, 1, 0, 0)]
    [InlineData("M.CC0001", 'M', 0, DocxTargetKind.ContentControl, 1, 0, 0)]
    [InlineData("M.S0001", 'M', 0, DocxTargetKind.Section, 1, 0, 0)]
    [InlineData("H001.P0002", 'H', 1, DocxTargetKind.Paragraph, 2, 0, 0)]
    [InlineData("F002.T0003", 'F', 2, DocxTargetKind.Table, 3, 0, 0)]
    [InlineData("H001.T0001.R02", 'H', 1, DocxTargetKind.Row, 1, 2, 0)]
    [InlineData("F001.T0001.R02.C03", 'F', 1, DocxTargetKind.Cell, 1, 2, 3)]
    [InlineData("H001.T0001.MG0001", 'H', 1, DocxTargetKind.MergeGroup, 1, 1, 0)]
    [InlineData("H001.I0004", 'H', 1, DocxTargetKind.Image, 4, 0, 0)]
    [InlineData("F001.L0002", 'F', 1, DocxTargetKind.Hyperlink, 2, 0, 0)]
    [InlineData("H001.F0003", 'H', 1, DocxTargetKind.Field, 3, 0, 0)]
    [InlineData("F001.B0004", 'F', 1, DocxTargetKind.Bookmark, 4, 0, 0)]
    [InlineData("H001.CC0005", 'H', 1, DocxTargetKind.ContentControl, 5, 0, 0)]
    [InlineData("M.T0001.R100", 'M', 0, DocxTargetKind.Row, 1, 100, 0)]
    [InlineData("M.P10000", 'M', 0, DocxTargetKind.Paragraph, 10000, 0, 0)]
    [InlineData("M.T10000.R02.C03", 'M', 0, DocxTargetKind.Cell, 10000, 2, 3)]
    [InlineData("M.T0001.R02.C100", 'M', 0, DocxTargetKind.Cell, 1, 2, 100)]
    [InlineData("M.T0001.MG10000", 'M', 0, DocxTargetKind.MergeGroup, 1, 10000, 0)]
    [InlineData("M.CC10000", 'M', 0, DocxTargetKind.ContentControl, 10000, 0, 0)]
    [InlineData("H1000.P0001", 'H', 1000, DocxTargetKind.Paragraph, 1, 0, 0)]
    public static void ParsesWireFormAndRoundTrips(
        string wire,
        char story,
        int storyPart,
        DocxTargetKind kind,
        int primary,
        int secondary,
        int tertiary)
    {
        Assert.True(DocxTargetId.TryParse(wire, out DocxTargetId targetId));
        Assert.Equal(new DocxTargetId(story, storyPart, kind, primary, secondary, tertiary), targetId);
        Assert.Equal(wire, targetId.ToWireValue());
        Assert.Equal(wire, targetId.ToString());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("M.P1")]
    [InlineData("M.P0001X")]
    [InlineData("X.P0001")]
    [InlineData("M.X0001")]
    [InlineData("m.p0001")]
    [InlineData("M.P000A")]
    [InlineData("H01.P0001")]
    [InlineData("M.T0001.R2.C03")]
    [InlineData("M.T0001.MG01")]
    [InlineData("M.CC001")]
    [InlineData("M.P-001")]
    [InlineData("M.P+001")]
    [InlineData("M.P 001")]
    [InlineData("M.P0001 ")]
    [InlineData(" M.P0001")]
    [InlineData("M.P0000")]
    [InlineData("M.T0001.R00")]
    [InlineData("M.T0001.R02.C00")]
    [InlineData("M.T0001.R-02")]
    [InlineData("H000.P0001")]
    [InlineData("M.R02")]
    [InlineData("M.MG0001")]
    [InlineData("M.T0001.C03")]
    [InlineData("M.P99999999999999999999")]
    public static void RejectsMalformedWireForm(string? wire)
    {
        Assert.False(DocxTargetId.TryParse(wire, out DocxTargetId _));
    }
}
