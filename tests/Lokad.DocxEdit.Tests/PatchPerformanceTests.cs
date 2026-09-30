using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// C04: allocation regression for the repeated merge-group walks. An empty check
// over a table with one horizontal merge per row once allocated ~94.5 MiB for
// 1,200 rows; the single-pass enumeration brought it to ~2.2 MiB. The bound
// below keeps 9x headroom over the fixed cost while staying far under the old
// quadratic blowup, so it pins the traversal shape without asserting wall time.
public static class PatchPerformanceTests
{
    [Fact]
    public static void EmptyPatchCheckOnLargeMergedTableStaysBounded()
    {
        var body = new StringBuilder("<w:tbl><w:tblGrid>");
        body.Append("<w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>");
        for (int row = 0; row < 1200; row++)
        {
            body.Append("<w:tr><w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr>");
            body.Append("<w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>");
        }

        body.Append("</w:tbl>");
        string bodyXml = body.ToString();
        using (MemoryStream warmup = CreateDocxWithBody(bodyXml))
        {
            Assert.True(new DocxEditor().Check(warmup, new StringReader("docxpatch 1\n")).Success);
        }

        using MemoryStream input = CreateDocxWithBody(bodyXml);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        DocxCheckResult result = new DocxEditor().Check(input, new StringReader("docxpatch 1\n"));
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        Assert.True(result.Success);
        Assert.True(allocated < 20L * 1024 * 1024, "Allocated " + allocated + " bytes.");
    }
}
