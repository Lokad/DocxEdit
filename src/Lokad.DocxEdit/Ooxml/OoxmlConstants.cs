using System.Xml.Linq;

namespace Lokad.DocxEdit.Ooxml;

internal static class OoxmlNs
{
    public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    public static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    public static readonly XNamespace W16Cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
    public static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    public static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    public static readonly XNamespace Xml = "http://www.w3.org/XML/1998/namespace";
}

internal static class OoxmlRelTypes
{
    public const string OfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    public const string Header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    public const string Footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    public const string Image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    public const string Styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    public const string Numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    public const string Settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    public const string Comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    public const string CommentsExtended = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
    public const string CommentsIds = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
    public const string Hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    public const string Footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    public const string Endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
}

internal static class OoxmlContentTypeNames
{
    public const string Xml = "application/xml";
    public const string MainDocument = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    public const string MacroEnabledMainDocument = "application/vnd.ms-word.document.macroEnabled.main+xml";
    public const string Relationships = "application/vnd.openxmlformats-package.relationships+xml";
}
