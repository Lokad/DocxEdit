using System.Xml;
using System.Xml.Linq;

namespace Lokad.DocxEdit.Ooxml;

internal static class SafeXml
{
    public static XmlReaderSettings CreateReaderSettings()
    {
        return new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreWhitespace = false,
            IgnoreComments = false
        };
    }

    public static XDocument Load(Stream stream, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using XmlReader reader = XmlReader.Create(stream, CreateReaderSettings());
        XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        cancellationToken.ThrowIfCancellationRequested();
        return document;
    }
}

