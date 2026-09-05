namespace Lokad.DocxEdit.Tests;

public static class DocxTargetIdTests
{
    // kindValue follows DocxTargetKind declaration order.
    [Theory]
    [InlineData("M.P0001", 'M', 0, 0, 1, 0, 0)]
    [InlineData("M.T0001", 'M', 0, 1, 1, 0, 0)]
    [InlineData("M.T0001.R02", 'M', 0, 2, 1, 2, 0)]
    [InlineData("M.T0001.R02.C03", 'M', 0, 3, 1, 2, 3)]
    [InlineData("M.T0001.MG0001", 'M', 0, 4, 1, 1, 0)]
    [InlineData("M.I0001", 'M', 0, 5, 1, 0, 0)]
    [InlineData("M.L0001", 'M', 0, 6, 1, 0, 0)]
    [InlineData("M.F0001", 'M', 0, 7, 1, 0, 0)]
    [InlineData("M.B0001", 'M', 0, 8, 1, 0, 0)]
    [InlineData("M.CC0001", 'M', 0, 9, 1, 0, 0)]
    [InlineData("M.S0001", 'M', 0, 10, 1, 0, 0)]
    [InlineData("H001.P0002", 'H', 1, 0, 2, 0, 0)]
    [InlineData("F002.T0003", 'F', 2, 1, 3, 0, 0)]
    [InlineData("H001.T0001.R02", 'H', 1, 2, 1, 2, 0)]
    [InlineData("F001.T0001.R02.C03", 'F', 1, 3, 1, 2, 3)]
    [InlineData("H001.T0001.MG0001", 'H', 1, 4, 1, 1, 0)]
    [InlineData("H001.I0004", 'H', 1, 5, 4, 0, 0)]
    [InlineData("F001.L0002", 'F', 1, 6, 2, 0, 0)]
    [InlineData("H001.F0003", 'H', 1, 7, 3, 0, 0)]
    [InlineData("F001.B0004", 'F', 1, 8, 4, 0, 0)]
    [InlineData("H001.CC0005", 'H', 1, 9, 5, 0, 0)]
    public static void ParsesWireFormAndRoundTrips(
        string wire,
        char story,
        int storyPart,
        int kindValue,
        int primary,
        int secondary,
        int tertiary)
    {
        Assert.True(DocxTargetId.TryParse(wire, out DocxTargetId targetId));
        Assert.Equal(new DocxTargetId(story, storyPart, (DocxTargetKind)kindValue, primary, secondary, tertiary), targetId);
        Assert.Equal(wire, targetId.ToWireValue());
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
    public static void RejectsMalformedWireForm(string? wire)
    {
        Assert.False(DocxTargetId.TryParse(wire, out DocxTargetId _));
    }
}
