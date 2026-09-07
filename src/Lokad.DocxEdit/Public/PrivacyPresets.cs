namespace Lokad.DocxEdit;

/// <summary>
/// Ready-made low-text options and renderers for privacy-sensitive runs: they keep
/// counts, IDs, and markup metadata while dropping document text.
/// </summary>
public static class DocxPrivacyPresets
{
    /// <summary>Read options without text: counts, IDs, and diagnostics only.</summary>
    public static DocxReadOptions ReadSummary => new()
    {
        MaxText = 0
    };

    /// <summary>Context options with neighbors but no text.</summary>
    public static DocxContextOptions ContextMetadataOnly => new()
    {
        Radius = 1,
        MaxText = 0
    };

    /// <summary>Changes options with markup metadata but no comment text.</summary>
    public static DocxChangesOptions ChangesMarkupOnly => new();

    /// <summary>Renders a read summary without document text.</summary>
    public static string RenderReadSummary(DocxReadResult result)
    {
        return DocxTextRenderer.RenderReadSummary(result);
    }

    /// <summary>Renders target context metadata without document text, even when the result carries text.</summary>
    public static string RenderContextMetadata(DocxContextResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        IReadOnlyList<DocxContextItem> redacted = result.Items
            .Select(item => item with { Text = string.Empty })
            .ToArray();
        return Rendering.TextRenderers.RenderContext(redacted);
    }

    /// <summary>Renders change markup metadata without comment text, even when the result carries snippets.</summary>
    public static string RenderChangesMarkup(DocxChangesResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        var redacted = result with
        {
            Changes = result.Changes
                .Select(change => change with { CommentTextSnippet = null, CommentTextTruncated = false })
                .ToArray(),
            CommentSummary = result.CommentSummary
                .Select(summary => summary with { TextSnippet = null, TextTruncated = false })
                .ToArray()
        };
        return DocxTextRenderer.RenderChanges(redacted);
    }
}
