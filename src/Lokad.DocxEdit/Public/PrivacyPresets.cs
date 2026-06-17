namespace Lokad.DocxEdit;

public static class DocxPrivacyPresets
{
    public static DocxReadOptions ReadSummary => new()
    {
        MaxText = 0
    };

    public static DocxContextOptions ContextMetadataOnly => new()
    {
        Radius = 1,
        MaxText = 0
    };

    public static DocxChangesOptions ChangesMarkupOnly => new();

    public static string RenderReadSummary(DocxReadResult result)
    {
        return DocxTextRenderer.RenderReadSummary(result);
    }

    public static string RenderContextMetadata(DocxContextResult result)
    {
        return DocxTextRenderer.RenderContext(result);
    }

    public static string RenderChangesMarkup(DocxChangesResult result)
    {
        return DocxTextRenderer.RenderChanges(result);
    }
}
