using DocxEdit.Ooxml;

namespace DocxEdit.Tests;

public static class OoxmlUtilityTests
{
    [Fact]
    public static void UnitConversionsReturnExpectedEmus()
    {
        Assert.Equal(914400, OoxmlUnits.InchesToEmu(1));
        Assert.Equal(360000, OoxmlUnits.CentimetersToEmu(1));
        Assert.Equal(12700, OoxmlUnits.PointsToEmu(1));
        Assert.Equal(914400, OoxmlUnits.PixelsToEmu(96));
    }

    [Theory]
    [InlineData("1in", 914400)]
    [InlineData("2.54cm", 914400)]
    [InlineData("72pt", 914400)]
    [InlineData("96px", 914400)]
    [InlineData("123emu", 123)]
    public static void TryParseDimensionSupportsDocumentedUnits(string text, long expectedEmus)
    {
        Assert.True(OoxmlUnits.TryParseDimension(text, out long emus));
        Assert.Equal(expectedEmus, emus);
    }

    [Fact]
    public static void AllocateRelationshipIdReturnsFirstAvailableRId()
    {
        string id = OoxmlIds.AllocateRelationshipId(["rId1", "rId2", "custom"]);

        Assert.Equal("rId3", id);
    }

    [Fact]
    public static void AllocateImagePartNameReturnsFirstAvailableMediaName()
    {
        string partName = OoxmlMediaParts.AllocateImagePartName(
            ["/word/media/image1.png", "/word/media/image2.png"],
            "image/png");

        Assert.Equal("/word/media/image3.png", partName);
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData("1", true)]
    [InlineData("true", true)]
    [InlineData("on", true)]
    [InlineData("0", false)]
    [InlineData("false", false)]
    public static void WmlBooleanParsesCommonWordprocessingValues(string? value, bool expected)
    {
        Assert.Equal(expected, WmlBoolean.IsTrue(value));
    }
}
