using System.IO.Compression;
using System.Xml.Linq;

namespace Lokad.DocxEdit.Tests;

public static class EquationTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    [Theory]
    [InlineData(@"\frac{-b\pm\sqrt{b^2-4ac}}{2a}", "f")]
    [InlineData(@"x_i^2+\alpha", "sSubSup")]
    [InlineData(@"\sqrt[3]{x}", "rad")]
    [InlineData(@"\sum_{i=1}^{n}{i^2}", "nary")]
    [InlineData(@"\int_0^1{\frac{1}{1+x^2}}\,\mathrm{d}x", "nary")]
    [InlineData(@"\begin{pmatrix}a&b\\c&d\end{pmatrix}", "m")]
    [InlineData(@"\left\langle\frac{x}{y}\right\rangle", "d")]
    [InlineData(@"\text{when }x\geq0", "r")]
    public static void CreatesNativeMath(string latex, string construct)
    {
        using var input = Empty();
        using var output = Apply(input, $"op insert-equation\ntarget M.P0001\nlatex {latex}\nend");
        DocxReadResult read = Read(output);
        DocxEquationInfo equation = Assert.Single(read.Equations);
        Assert.Equal("M.E0001", equation.Id.ToWireValue());
        Assert.Equal("M.P0002", equation.TargetId?.ToWireValue());
        Assert.True(equation.IsDisplay);
        XElement xml = XElement.Parse(equation.Omml!);
        Assert.Contains(xml.Descendants(), e => e.Name == M + construct);
        Assert.DoesNotContain("snapshot", equation.Omml);
        Assert.Equal(64, equation.ContentHash.Length);
        output.Position = 0;
        Assert.True(new DocxEditor().Validate(output).Success);
    }

    [Theory]
    [InlineData(@"\unknown{x}")]
    [InlineData(@"\frac{x}")]
    [InlineData(@"x^2^3")]
    [InlineData(@"\sqrt[]{x}")]
    [InlineData(@"\begin{matrix}a&b\\c\end{matrix}")]
    [InlineData(@"\begin{matrix}a\end{pmatrix}")]
    [InlineData(@"\left(x")]
    [InlineData(@"$x$")]
    [InlineData(@"\sum_0^n x")]
    public static void InvalidSourceFailsCheckAndApplyWithoutOutput(string latex)
    {
        using var input = Empty();
        string patch = $"docxpatch 1\nop insert-equation\ntarget M.P0001\nlatex {latex}\nend";
        var editor = new DocxEditor();
        var check = editor.Check(input, new StringReader(patch));
        Assert.False(check.Success);
        input.Position = 0;
        using var output = new MemoryStream();
        var apply = editor.Apply(input, new StringReader(patch), output);
        Assert.False(apply.Success);
        Assert.Equal(0, output.Length);
        Assert.Contains(apply.Diagnostics, d => d.Code == "E4205" && d.Line == 4);
        Assert.Equal(check.Diagnostics.Select(d => d.Code), apply.Diagnostics.Select(d => d.Code));
    }

    [Fact]
    public static void SnapshotGuardsAliasesAndFinalIdsSurviveEdits()
    {
        using var input = Synthetic(new XElement(W + "p", Math("a"), Math("b")));
        DocxEquationInfo second = Read(input).Equations[1];
        input.Position = 0;
        using var output = Apply(input, $"""
            op delete-equation
            target M.E0001
            end
            op replace-equation
            target M.E0002
            expect-hash {second.ContentHash}
            latex z^2
            end
            op insert-equation
            target M.P0001
            placement inline
            latex x
            as added
            end
            op replace-equation
            target @added
            latex y
            end
            """, out var result);
        Assert.Equal(new[] { "z2", "y" }, Read(output).Equations.Select(e => e.Text));
        Assert.Contains(result.Operations, o => o.CreatedTargetIds.Contains("M.E0002"));
        var deletion = Assert.Single(result.Operations[0].AffectedTargets);
        Assert.Equal("delete", deletion.Action);
        Assert.Null(deletion.FinalId);
        var replaced = Assert.Single(result.Operations[1].AffectedTargets);
        Assert.Equal("M.E0002", replaced.Id.ToWireValue());
        Assert.Equal("M.E0001", replaced.FinalId?.ToWireValue());
        var aliased = Assert.Single(result.Operations[3].AffectedTargets);
        Assert.Equal("equation", aliased.Kind);
        Assert.Equal("M.E0002", aliased.FinalId?.ToWireValue());
        Assert.DoesNotContain("snapshot", DocxTestFixtures.ReadEntry(output, "word/document.xml"));
        Assert.DoesNotContain("alias=", DocxTestFixtures.ReadEntry(output, "word/document.xml"));
    }

    [Fact]
    public static void StaleGuardAndDeletedIdFailWithoutRetargeting()
    {
        using var input = Synthetic(new XElement(W + "p", Math("a"), Math("b")));
        AssertFails(input, "op replace-equation\ntarget M.E0001\nexpect-hash stale\nlatex z\nend", "E3201");
        AssertFails(input, "op delete-equation\ntarget M.E0001\nend\nop delete-equation\ntarget M.E0001\nend", "E1201");
    }

    [Fact]
    public static void SurroundingTextIsPreservedAndCrossEquationMatchIsRejected()
    {
        using var input = Synthetic(new XElement(W + "p", Run("before"), Math("x"), Run("after")));
        AssertFails(input, "op replace-paragraph\ntarget M.P0001\ntext erased\nend", "E4305");
        AssertFails(input, "op replace-text\ntarget M.P0001\nfind beforeafter\nwith erased\nend", "E4305");
        input.Position = 0;
        using var edited = Apply(input, "op replace-text\ntarget M.P0001\nfind before\nwith BEFORE\nend");
        Assert.Equal("BEFOREafter", Read(edited).Paragraphs[0].Text);
        Assert.Single(Read(edited).Equations);
        edited.Position = 0;
        using var removed = Apply(edited, "op delete-equation\ntarget M.E0001\nend");
        Assert.Empty(Read(removed).Equations);
        Assert.Equal("BEFOREafter", Read(removed).Paragraphs[0].Text);
    }

    [Theory]
    [InlineData("cell", "set-cell", "M.T0001.R01.C01", "force true")]
    [InlineData("hyperlink", "set-hyperlink-text", "M.L0001", "")]
    [InlineData("field", "set-field-result", "M.F0001", "")]
    [InlineData("control", "set-content-control-text", "M.CC0001", "")]
    public static void ContainerRewritesCannotEraseEquations(string shape, string operation, string target, string extra)
    {
        XElement content = shape switch
        {
            "cell" => new XElement(W + "tbl", new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", "2000"))),
                new XElement(W + "tr", new XElement(W + "tc", new XElement(W + "p", Math("x"))))),
            "hyperlink" => new XElement(W + "p", new XElement(W + "hyperlink", Math("x"))),
            "field" => new XElement(W + "p", new XElement(W + "fldSimple", new XAttribute(W + "instr", "QUOTE 1"), Math("x"))),
            _ => new XElement(W + "sdt", new XElement(W + "sdtPr", new XElement(W + "text")), new XElement(W + "sdtContent", new XElement(W + "p", Math("x"))))
        };
        using var input = Synthetic(content);
        AssertFails(input, $"op {operation}\ntarget {target}\ntext removed\n{extra}\nend", "E4305");
        input.Position = 0;
        var capabilities = new DocxEditor().GetCapabilities(input, target);
        Assert.Equal("unsupported", Assert.Single(capabilities.Capabilities!.Operations, o => o.Operation == operation).Support);
    }

    [Fact]
    public static void EquationInspectionRespectsPrivacyBounds()
    {
        using var input = Synthetic(new XElement(W + "p", Math("sensitive")));
        var editor = new DocxEditor();
        var read = editor.Read(input, new DocxReadOptions { MaxText = 0 });
        Assert.Empty(Assert.Single(read.Equations).Text);
        Assert.Null(read.Equations[0].Omml);
        input.Position = 0;
        var dump = editor.Dump(input, "M.E0001", new DocxDumpOptions { MaxText = 0 });
        Assert.True(dump.Success);
        Assert.Null(dump.Equation!.Omml);
        Assert.DoesNotContain("sensitive", System.Text.Json.JsonSerializer.Serialize(dump));
        input.Position = 0;
        Assert.True(editor.Context(input, "M.E0001").Success);
        input.Position = 0;
        Assert.Contains(editor.Outline(input).Items, i => i.Kind == "equation");
    }

    [Fact]
    public static void ProtectedNativeEquationIsInspectableButCannotBeRewritten()
    {
        using var input = Synthetic(new XElement(W + "p", new XElement(M + "oMath", new XElement(W + "bookmarkStart", new XAttribute(W + "id", "1"), new XAttribute(W + "name", "keep")),
            new XElement(M + "r", new XElement(M + "t", "x")), new XElement(W + "bookmarkEnd", new XAttribute(W + "id", "1")))));
        Assert.Single(Read(input).Equations);
        AssertFails(input, "op replace-equation\ntarget M.E0001\nlatex z\nend", "E4305");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public static void FormattingRevisionsAndComplexFieldResultsAreProtected(bool formattingRevision)
    {
        XElement math = Math("x");
        XElement paragraph = new(W + "p", math);
        if (formattingRevision)
        {
            math.Element(M + "r")!.AddFirst(new XElement(W + "rPr", new XElement(W + "rPrChange",
                new XAttribute(W + "id", "1"), new XAttribute(W + "author", "Test"), new XElement(W + "rPr"))));
        }
        else
        {
            math.AddBeforeSelf(new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
                new XElement(W + "r", new XElement(W + "instrText", "QUOTE 1")),
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))));
            math.AddAfterSelf(new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));
        }
        using var input = Synthetic(paragraph);
        AssertFails(input, "op replace-equation\ntarget M.E0001\nlatex y\nend", "E4305");
        input.Position = 0;
        Assert.All(new DocxEditor().GetCapabilities(input, "M.E0001").Capabilities!.Operations, o => Assert.Equal("unsupported", o.Support));
    }

    [Theory]
    [InlineData(TrackChangesMode.Off, true)]
    [InlineData(TrackChangesMode.Suggest, true)]
    [InlineData(TrackChangesMode.Require, false)]
    public static void TrackingPolicyIsExplicit(TrackChangesMode mode, bool success)
    {
        using var input = Synthetic(new XElement(W + "p", Math("x")));
        using var output = new MemoryStream();
        var result = new DocxEditor().Apply(input, new StringReader("docxpatch 1\nop replace-equation\ntarget M.E0001\nlatex y\nend"), output, new DocxEditOptions { TrackChanges = mode });
        Assert.Equal(success, result.Success);
        if (mode == TrackChangesMode.Suggest) Assert.Contains(result.Diagnostics, d => d.Code == "W4001");
        if (!success) Assert.Contains(result.Diagnostics, d => d.Code == "E6001");
    }

    [Fact]
    public static void ResourceLimitsRejectPathologicalInput()
    {
        using var input = Empty();
        AssertFails(input, "op insert-equation\ntarget M.P0001\nlatex " + new string('{', 1000) + "x" + new string('}', 1000) + "\nend", "E4205");
        AssertFails(input, "op insert-equation\ntarget M.P0001\nlatex " + new string('x', 32769) + "\nend", "E4205");
        AssertFails(input, "op insert-equation\ntarget M.P0001\nlatex " + string.Concat(Enumerable.Repeat(@"\sum_", 1000)) + "x\nend", "E4205");
    }

    [Fact]
    public static void DisplayGroupAndParagraphPropertiesSurviveWholeEquationEdits()
    {
        using var input = Synthetic(new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", "right"))),
            new XElement(M + "oMathPara", new XElement(M + "oMathParaPr", new XElement(M + "jc", new XAttribute(M + "val", "right"))), Math("a"), Math("b"))));
        string survivingHash = Read(input).Equations[1].ContentHash;
        using var deleted = Apply(input, "op delete-equation\ntarget M.E0001\nend");
        Assert.Equal(survivingHash, Assert.Single(Read(deleted).Equations).ContentHash);
        string xml = DocxTestFixtures.ReadEntry(deleted, "word/document.xml");
        Assert.Single(XDocument.Parse(xml).Descendants(M + "oMathParaPr"));
        using var replaced = Apply(deleted, "op replace-equation\ntarget M.E0001\nlatex b^2\nend");
        Assert.True(Assert.Single(Read(replaced).Equations).IsDisplay);
        using var empty = Apply(replaced, "op delete-equation\ntarget M.E0001\nend");
        var final = XDocument.Parse(DocxTestFixtures.ReadEntry(empty, "word/document.xml"));
        Assert.Empty(final.Descendants(M + "oMathPara"));
        Assert.Equal("right", (string?)Assert.Single(final.Descendants(W + "jc")).Attribute(W + "val"));
    }

    [Fact]
    public static void PhysicalIdsRemainStableAcrossRevisionViews()
    {
        using var input = Synthetic(new XElement(W + "p", new XElement(W + "del", new XAttribute(W + "id", "1"), new XAttribute(W + "author", "Test"), Math("old")),
            Math("unchanged"), new XElement(W + "ins", new XAttribute(W + "id", "2"), new XAttribute(W + "author", "Test"), Math("new"))));
        var editor = new DocxEditor();
        Assert.Equal(new[] { "M.E0002", "M.E0003" }, editor.Read(input).Equations.Select(e => e.Id.ToWireValue()));
        input.Position = 0;
        Assert.Equal(new[] { "M.E0001", "M.E0002" }, editor.Read(input, new DocxReadOptions { TextView = DocxTextView.Original }).Equations.Select(e => e.Id.ToWireValue()));
        AssertFails(input, "op replace-equation\ntarget M.E0003\nlatex z\nend", "E4305");
    }

    [Fact]
    public static void EquationCapabilitiesAndTemplatesRespectPolicy()
    {
        using var input = Synthetic(new XElement(W + "p", Math("x")));
        var editor = new DocxEditor();
        var capabilities = editor.GetCapabilities(input, "M.E0001");
        Assert.True(capabilities.Success);
        Assert.All(capabilities.Capabilities!.Operations, c => Assert.Equal("supported", c.Support));
        input.Position = 0;
        var template = editor.GetTemplate(input, "M.E0001");
        Assert.True(template.Success);
        Assert.Contains("# op replace-equation", template.Template);
        Assert.Contains("expect-hash " + Read(input).Equations[0].ContentHash, template.Template);
        input.Position = 0;
        Assert.All(editor.GetCapabilities(input, "M.E0001", new DocxCapabilitiesOptions { TrackChanges = TrackChangesMode.Require }).Capabilities!.Operations,
            c => Assert.Equal("unsupported", c.Support));
    }

    [Fact]
    public static void HeaderAndFooterEquationsCanBeDiscoveredAndEdited()
    {
        using var input = Empty();
        XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        using (var zip = new ZipArchive(input, ZipArchiveMode.Update, true))
        {
            foreach (var (name, root) in new[] { ("header", "hdr"), ("footer", "ftr") })
            {
                Update("word/_rels/document.xml.rels", d => d.Root!.Add(new XElement(relationships + "Relationship", new XAttribute("Id", name),
                    new XAttribute("Type", r.NamespaceName + "/" + name), new XAttribute("Target", name + ".xml"))));
                Update("[Content_Types].xml", d => d.Root!.Add(new XElement(types + "Override", new XAttribute("PartName", "/word/" + name + ".xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml." + name + "+xml"))));
                Update("word/document.xml", d => d.Descendants(W + "sectPr").Single().AddFirst(new XElement(W + (name + "Reference"),
                    new XAttribute(W + "type", "default"), new XAttribute(r + "id", name))));
                DocxTestFixtures.AddEntry(zip, "word/" + name + ".xml", new XElement(W + root, new XElement(W + "p", Math(name))).ToString());
            }
            void Update(string partName, Action<XDocument> edit)
            {
                var entry = zip.GetEntry(partName)!;
                XDocument document;
                using (var stream = entry.Open()) document = XDocument.Load(stream);
                edit(document);
                entry.Delete();
                using var output = zip.CreateEntry(partName).Open();
                document.Save(output);
            }
        }
        input.Position = 0;
        var editor = new DocxEditor();
        Assert.Empty(editor.Read(input).Equations);
        input.Position = 0;
        Assert.Equal(new[] { "H001.E0001", "F001.E0001" }, editor.Read(input, new DocxReadOptions { IncludeHeadersFooters = true }).Equations.Select(e => e.Id.ToWireValue()));
        using var edited = Apply(input, "op replace-equation\ntarget H001.E0001\nlatex h^2\nend\nop delete-equation\ntarget F001.E0001\nend");
        var read = editor.Read(edited, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Equal("h2", Assert.Single(read.Equations).Text);
        Assert.Equal("H001.P0001", read.Equations[0].TargetId?.ToWireValue());
    }

    private static MemoryStream Empty()
    {
        var stream = new MemoryStream();
        Assert.True(new DocxEditor().Create(stream).Success);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream Synthetic(XElement block)
    {
        var stream = Empty();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true))
        {
            var entry = zip.GetEntry("word/document.xml")!;
            XDocument document;
            using (var part = entry.Open()) document = XDocument.Load(part);
            document.Descendants(W + "p").First().ReplaceWith(block);
            entry.Delete();
            using var output = zip.CreateEntry("word/document.xml").Open();
            document.Save(output);
        }
        stream.Position = 0;
        return stream;
    }

    private static XElement Math(string text) => new(M + "oMath", new XElement(M + "r", new XElement(M + "t", text)));
    private static XElement Run(string text) => new(W + "r", new XElement(W + "t", text));
    private static DocxReadResult Read(MemoryStream stream) { stream.Position = 0; return new DocxEditor().Read(stream, new DocxReadOptions { MaxText = 100000 }); }
    private static MemoryStream Apply(MemoryStream input, string operations) => Apply(input, operations, out _);
    private static MemoryStream Apply(MemoryStream input, string operations, out DocxApplyResult result)
    {
        input.Position = 0;
        var output = new MemoryStream();
        result = new DocxEditor().Apply(input, new StringReader("docxpatch 1\n" + operations), output);
        Assert.True(result.Success, string.Join("; ", result.Diagnostics));
        output.Position = 0;
        return output;
    }
    private static void AssertFails(MemoryStream input, string operations, string code)
    {
        input.Position = 0;
        using var output = new MemoryStream();
        var result = new DocxEditor().Apply(input, new StringReader("docxpatch 1\n" + operations), output);
        Assert.False(result.Success);
        Assert.Equal(0, output.Length);
        Assert.Contains(result.Diagnostics, d => d.Code == code);
    }
}
