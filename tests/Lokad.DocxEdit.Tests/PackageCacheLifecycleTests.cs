using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Xml.Linq;

namespace Lokad.DocxEdit.Tests;

// Exercise mutations unavailable through the public patch language, without
// extending the library's integration surface just to expose internal caches.
public static class PackageCacheLifecycleTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string StylesType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

    [Fact]
    public static void StyleResolutionFollowsReplacementRemovalAndRelocation()
    {
        object package = Blank();
        Assert.Equal(1, Heading(package));
        Call(package, "ReplacePartBytes", "/word/styles.xml", Styles(2));
        Assert.Equal(3, Heading(package));
        Call(package, "RemovePart", "/word/styles.xml", CancellationToken.None);
        Assert.Equal(1, Heading(package));
        Call(package, "AddPart", "/word/styles.xml", StylesType, Styles(9), CancellationToken.None);
        Assert.Null(Heading(package));
        Call(package, "AddPart", "/custom/styles.xml", StylesType, Styles(1), CancellationToken.None);
        Call(package, "ReplacePartBytes", "/word/_rels/document.xml.rels", Encoding.UTF8.GetBytes(
            $"<Relationships xmlns=\"{Rel}\"><Relationship Id=\"rStyles\" Type=\"{R}/styles\" Target=\"../custom/styles.xml\"/></Relationships>"));
        Assert.Equal(2, Heading(package));
    }

    [Fact]
    public static void RelationshipListsFollowAddRemoveAndRecreationAndHonorCancellation()
    {
        object package = Blank();
        Assert.Single(Relationships(package, "/word/document.xml"));
        Call(package, "AddRelationship", "/word/document.xml", "rLink", R + "/hyperlink", "https://example.test", "External", CancellationToken.None);
        Assert.Equal(2, Relationships(package, "/word/document.xml").Count);
        Call(package, "RemoveRelationship", "/word/document.xml", "rLink", CancellationToken.None);
        Assert.Single(Relationships(package, "/word/document.xml"));
        Assert.Empty(Relationships(package, "/word/header1.xml"));
        Call(package, "AddRelationship", "/word/header1.xml", "rLink", R + "/hyperlink", "https://example.test", "External", CancellationToken.None);
        Assert.Single(Relationships(package, "/word/header1.xml"));
        Call(package, "RemovePart", "/word/_rels/header1.xml.rels", CancellationToken.None);
        Assert.Empty(Relationships(package, "/word/header1.xml"));
        Call(package, "AddPart", "/word/_rels/header1.xml.rels", "application/vnd.openxmlformats-package.relationships+xml",
            Encoding.UTF8.GetBytes($"<Relationships xmlns=\"{Rel}\"><Relationship Id=\"replacement\" Type=\"{R}/hyperlink\" Target=\"https://example.test/new\" TargetMode=\"External\"/></Relationships>"), CancellationToken.None);
        object relationship = Assert.Single(Relationships(package, "/word/header1.xml").Cast<object>());
        Assert.Equal("replacement", relationship.GetType().GetProperty("Id")!.GetValue(relationship));
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => Call(package, "GetRelationships", "/word/document.xml", canceled.Token));
        Assert.Throws<OperationCanceledException>(() => Call(package, "GetPartRoleIndex", canceled.Token));
        Assert.Throws<OperationCanceledException>(() => Call(package, "GetStyleIndex", canceled.Token));
    }

    [Fact]
    public static void RemovedAndRecreatedXmlDoesNotReturnTheOldParsedTree()
    {
        object package = Blank();
        Type engine = typeof(DocxEditor).Assembly.GetType("Lokad.DocxEdit.DocxPatchEngine", throwOnError: true)!;
        object?[] args = [package, "/word/document.xml", CancellationToken.None, null];
        _ = Call(engine, "LoadDocumentPart", args);
        Call(package, "RemovePart", "/word/document.xml", CancellationToken.None);
        Call(package, "AddPart", "/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            Encoding.UTF8.GetBytes($"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:r><w:t>Recreated</w:t></w:r></w:p></w:body></w:document>"), CancellationToken.None);
        XDocument document = (XDocument)Call(engine, "LoadDocumentPart", args)!;
        Assert.Equal("Recreated", Assert.Single(document.Descendants(XName.Get("t", W))).Value);
    }

    private static int? Heading(object package)
    {
        object index = Call(package, "GetStyleIndex", CancellationToken.None)!;
        return (int?)Call(index, "GetHeadingLevel", "Heading1", CancellationToken.None);
    }
    private static System.Collections.IList Relationships(object package, string name) =>
        (System.Collections.IList)Call(package, "GetRelationships", name, CancellationToken.None)!;
    private static byte[] Styles(int outline) => Encoding.UTF8.GetBytes(
        $"<w:styles xmlns:w=\"{W}\"><w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:pPr><w:outlineLvl w:val=\"{outline}\"/></w:pPr></w:style></w:styles>");
    private static object Blank()
    {
        Type package = typeof(DocxEditor).Assembly.GetType("Lokad.DocxEdit.Ooxml.OoxmlPackage", throwOnError: true)!;
        return Call(package, "CreateBlank", DocxPaperSize.A4, DocxOrientation.Portrait, DocxPackageLimits.Default, CancellationToken.None)!;
    }
    private static object? Call(object target, string name, params object?[] args)
    {
        Type type = target as Type ?? target.GetType();
        MethodInfo method = type.GetMethod(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static)!;
        Assert.NotNull(method);
        try { return method.Invoke(target is Type ? null : target, args); }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
            throw;
        }
    }
}
