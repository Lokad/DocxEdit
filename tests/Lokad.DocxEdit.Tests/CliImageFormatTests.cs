using System.Text;
using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

[Collection("ConsoleCli")]
public static class CliImageFormatTests
{
    [Fact]
    public static void SvgFileAssetRoundTripsThroughCli() => AssertRoundTrip(SvgImageTests.Diagram(), "svg");

    internal static void AssertRoundTrip(byte[] bytes, string extension)
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        string patch = Path.Combine(temp.Path, "edit.patch");
        string output = Path.Combine(temp.Path, "out.docx");
        string export = Path.Combine(temp.Path, "exports");
        using (var doc = ImageFormatTestTools.NewDocument()) File.WriteAllBytes(input, doc.ToArray());
        File.WriteAllBytes(Path.Combine(temp.Path, "asset." + extension), bytes);
        File.WriteAllText(patch, $"docxpatch 1\nop insert-image-after\ntarget M.P0001\nasset asset.{extension}\nend\n", Encoding.UTF8);
        Assert.Equal(0, CaptureCliOutput(null, ["check", input, patch]).ExitCode);
        Assert.Equal(0, CaptureCliOutput(null, ["apply", input, patch, "-o", output]).ExitCode);
        Assert.Equal(0, CaptureCliOutput(null, ["media", output, "--extract", export]).ExitCode);
        Assert.Equal(bytes, File.ReadAllBytes(Assert.Single(Directory.GetFiles(export, "*." + extension))));
    }
}
