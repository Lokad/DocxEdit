using System.Xml.Linq;

namespace Lokad.DocxEdit.Ooxml;

internal sealed partial class OoxmlPackage
{
    internal static OoxmlPackage CreateBlank(DocxPaperSize paperSize, DocxOrientation orientation, DocxPackageLimits quotas, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        (int width, int height) = paperSize switch
        {
            DocxPaperSize.A4 => (11906, 16838),
            DocxPaperSize.Letter => (12240, 15840),
            _ => throw new ArgumentOutOfRangeException(nameof(paperSize))
        };
        if (orientation == DocxOrientation.Landscape) (width, height) = (height, width);

        const string main = "/word/document.xml";
        const string stylesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
        XNamespace w = OoxmlNs.W;
        XNamespace ct = OoxmlNs.Ct;
        XNamespace rel = OoxmlNs.Rel;
        var parts = new Dictionary<string, OoxmlPart>(StringComparer.OrdinalIgnoreCase);
        Add("/[Content_Types].xml", OoxmlContentTypeNames.Xml,
            new XElement(ct + "Types",
                new XElement(ct + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", OoxmlContentTypeNames.Relationships)),
                new XElement(ct + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", OoxmlContentTypeNames.Xml)),
                new XElement(ct + "Override", new XAttribute("PartName", main), new XAttribute("ContentType", OoxmlContentTypeNames.MainDocument)),
                new XElement(ct + "Override", new XAttribute("PartName", "/word/styles.xml"), new XAttribute("ContentType", stylesContentType))));
        Add("/_rels/.rels", OoxmlContentTypeNames.Relationships,
            new XElement(rel + "Relationships", Relationship("rDocument", OoxmlRelTypes.OfficeDocument, "word/document.xml")));
        Add("/word/_rels/document.xml.rels", OoxmlContentTypeNames.Relationships,
            new XElement(rel + "Relationships", Relationship("rStyles", OoxmlRelTypes.Styles, "styles.xml")));
        Add(main, OoxmlContentTypeNames.MainDocument,
            new XElement(w + "document", new XAttribute(XNamespace.Xmlns + "w", w),
                new XElement(w + "body",
                    new XElement(w + "p"),
                    new XElement(w + "sectPr",
                        new XElement(w + "pgSz", new XAttribute(w + "w", width), new XAttribute(w + "h", height), new XAttribute(w + "orient", orientation.ToWireValue())),
                        new XElement(w + "pgMar",
                            new XAttribute(w + "top", 1440), new XAttribute(w + "right", 1440),
                            new XAttribute(w + "bottom", 1440), new XAttribute(w + "left", 1440),
                            new XAttribute(w + "header", 720), new XAttribute(w + "footer", 720), new XAttribute(w + "gutter", 0)),
                        new XElement(w + "cols", new XAttribute(w + "num", 1))))));

        var styles = new XElement(w + "styles", new XAttribute(XNamespace.Xmlns + "w", w),
            new XElement(w + "docDefaults",
                new XElement(w + "rPrDefault", new XElement(w + "rPr",
                    new XElement(w + "rFonts", new XAttribute(w + "ascii", "Arial"), new XAttribute(w + "hAnsi", "Arial"), new XAttribute(w + "eastAsia", "Arial"), new XAttribute(w + "cs", "Arial")),
                    Val("sz", 22), Val("szCs", 22))),
                new XElement(w + "pPrDefault", new XElement(w + "pPr",
                    new XElement(w + "spacing", new XAttribute(w + "after", 160), new XAttribute(w + "line", 240), new XAttribute(w + "lineRule", "auto"))))),
            new XElement(w + "style", new XAttribute(w + "type", "paragraph"), new XAttribute(w + "default", 1), new XAttribute(w + "styleId", "Normal"),
                Val("name", "Normal"), new XElement(w + "qFormat")));
        for (int level = 1; level <= 9; level++)
        {
            int fontSize = level switch { 1 => 32, 2 => 28, 3 => 24, _ => 22 };
            styles.Add(new XElement(w + "style", new XAttribute(w + "type", "paragraph"), new XAttribute(w + "styleId", $"Heading{level}"),
                Val("name", $"heading {level}"), Val("basedOn", "Normal"), Val("next", "Normal"),
                new XElement(w + "qFormat"),
                new XElement(w + "pPr", new XElement(w + "keepNext"), new XElement(w + "keepLines"),
                    new XElement(w + "spacing", new XAttribute(w + "before", 240), new XAttribute(w + "after", 120)), Val("outlineLvl", level - 1)),
                new XElement(w + "rPr", new XElement(w + "b"), new XElement(w + "bCs"), Val("sz", fontSize), Val("szCs", fontSize))));
        }
        Add("/word/styles.xml", stylesContentType, styles);

        if (parts.Count > quotas.MaxZipEntries || parts.Values.Sum(part => (long)part.Bytes.Length) > quotas.MaxUncompressedBytes)
        {
            throw new InvalidDataException("Created OOXML package exceeds the configured package quotas.");
        }
        return new OoxmlPackage(parts, main);

        XElement Val(string name, object value) => new(w + name, new XAttribute(w + "val", value));
        XElement Relationship(string id, string type, string target) => new(rel + "Relationship", new XAttribute("Id", id), new XAttribute("Type", type), new XAttribute("Target", target));
        void Add(string name, string contentType, XElement root)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var buffer = new MemoryStream();
            new XDocument(root).Save(buffer, SaveOptions.DisableFormatting);
            if (buffer.Length > quotas.MaxSinglePartBytes)
            {
                throw new InvalidDataException($"Created OOXML part '{name}' exceeds the configured single-part quota.");
            }
            parts.Add(name, new OoxmlPart(name, name.TrimStart('/'), contentType, buffer.ToArray()));
        }
    }
}
