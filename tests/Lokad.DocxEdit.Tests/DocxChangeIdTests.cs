namespace Lokad.DocxEdit.Tests;

public static class DocxChangeIdTests
{
    [Theory]
    [InlineData("M.CH0001", 'M', 0, 1)]
    [InlineData("M.CH0042", 'M', 0, 42)]
    [InlineData("H001.CH0002", 'H', 1, 2)]
    [InlineData("F002.CH0003", 'F', 2, 3)]
    [InlineData("C001.CH0001", 'C', 1, 1)]
    [InlineData("P001.CH0007", 'P', 1, 7)]
    [InlineData("M.CH10000", 'M', 0, 10000)]
    [InlineData("H001.CH10000", 'H', 1, 10000)]
    [InlineData("H1000.CH0001", 'H', 1000, 1)]
    public static void ParsesWireFormAndRoundTrips(string wire, char story, int storyPart, int ordinal)
    {
        Assert.True(DocxChangeId.TryParse(wire, out DocxChangeId changeId));
        Assert.Equal(new DocxChangeId(story, storyPart, ordinal), changeId);
        Assert.Equal(wire, changeId.ToWireValue());
        Assert.Equal(wire, changeId.ToString());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("M.CH1")]
    [InlineData("M.CH0001X")]
    [InlineData("X.CH0001")]
    [InlineData("M.P0001")]
    [InlineData("m.ch0001")]
    [InlineData("M.CH000A")]
    [InlineData("H01.CH0001")]
    [InlineData("M.T0001.MG0001")]
    [InlineData("M.CH0000")]
    [InlineData("M.CH-001")]
    [InlineData("M.CH 001")]
    [InlineData("H000.CH0001")]
    [InlineData("M.CH0001 ")]
    public static void RejectsMalformedWireForm(string? wire)
    {
        Assert.False(DocxChangeId.TryParse(wire, out DocxChangeId _));
    }
}
