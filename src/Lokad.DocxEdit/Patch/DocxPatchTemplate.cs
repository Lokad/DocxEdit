using System.Text;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

// D06: guarded patch templates generated from D17 capabilities. The active
// block passes check under the requested policy and changes nothing visible;
// every other supported operation is emitted as a commented block. Verdicts
// come from GetTargetCapabilities, so templates cannot disagree with check.
internal static partial class DocxPatchEngine
{
    internal sealed class TargetTemplateOutcome
    {
        public string? Template { get; init; }
        public DocxTargetCapabilities? Capabilities { get; init; }
        public DocxDiagnostic? Error { get; init; }
    }

    internal static TargetTemplateOutcome GetTargetTemplate(
        OoxmlPackage package,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        if (!DocxTargetId.TryParse(requestedTargetId, out DocxTargetId parsed))
        {
            return new TargetTemplateOutcome { Error = ParagraphCapabilitiesNotFound(requestedTargetId).Error };
        }

        ParagraphCapabilitiesOutcome capabilitiesOutcome = GetParagraphCapabilities(package, requestedTargetId, mode, cancellationToken);
        if (capabilitiesOutcome.Capabilities is null)
        {
            return new TargetTemplateOutcome { Error = capabilitiesOutcome.Error };
        }

        StoryDocument? storyDocument = TryResolveStoryDocument(package, parsed, cancellationToken);
        if (storyDocument is null)
        {
            return new TargetTemplateOutcome { Error = ParagraphCapabilitiesNotFound(requestedTargetId).Error };
        }

        DocxTargetCapabilities capabilities = capabilitiesOutcome.Capabilities;
        string? template = parsed.Kind switch
        {
            DocxTargetKind.Paragraph => BuildParagraphTemplate(storyDocument, parsed, capabilities, mode),
            DocxTargetKind.ContentControl => BuildContentControlTemplate(storyDocument, parsed, capabilities, mode),
            DocxTargetKind.Cell => BuildCellTemplate(storyDocument, parsed, capabilities, mode, isMergeGroup: false),
            DocxTargetKind.MergeGroup => BuildCellTemplate(storyDocument, parsed, capabilities, mode, isMergeGroup: true),
            DocxTargetKind.Table => BuildTableTemplate(storyDocument, parsed, capabilities, mode),
            DocxTargetKind.Row => BuildRowTemplate(storyDocument, parsed, capabilities, mode),
            DocxTargetKind.Section => BuildSectionTemplate(storyDocument, parsed, capabilities, mode),
            DocxTargetKind.Hyperlink => BuildHyperlinkTemplate(storyDocument, parsed, capabilities, mode),
            DocxTargetKind.Field => BuildFieldTemplate(storyDocument, parsed, capabilities, mode),
            DocxTargetKind.Image => BuildImageTemplate(storyDocument, parsed, capabilities, mode, package, cancellationToken),
            DocxTargetKind.Bookmark => BuildBookmarkTemplate(storyDocument, parsed, capabilities, mode),
            _ => null,
        };

        if (template is null)
        {
            return new TargetTemplateOutcome { Error = ParagraphCapabilitiesNotFound(requestedTargetId).Error };
        }

        return new TargetTemplateOutcome { Template = template, Capabilities = capabilities };
    }

    private static string SupportOf(DocxTargetCapabilities capabilities, string operation)
    {
        return capabilities.Operations
            .First(candidate => string.Equals(candidate.Operation, operation, StringComparison.Ordinal))
            .Support;
    }

    private static string ReasonOf(DocxTargetCapabilities capabilities, string operation)
    {
        return capabilities.Operations
            .First(candidate => string.Equals(candidate.Operation, operation, StringComparison.Ordinal))
            .Reason;
    }

    private static string ModeWord(TrackChangesMode mode)
    {
        return mode.ToString().ToLowerInvariant();
    }

    private static void AppendTemplateHeader(StringBuilder builder, DocxTargetCapabilities capabilities, TrackChangesMode mode, bool hasActive)
    {
        builder.Append("# Guarded patch template for ").Append(capabilities.TargetId)
            .Append(" (").Append(capabilities.Kind).Append(", story ").Append(capabilities.Story)
            .Append(") under track-changes ").Append(ModeWord(mode)).AppendLine(".");
        if (hasActive)
        {
            builder.AppendLine("# The first block below passes check and changes nothing visible; edit its text block to make a real change.");
        }
        else
        {
            builder.AppendLine("# No check-clean starter exists under this policy; uncomment a block below and fill its values, then check.");
        }

        builder.AppendLine("# Remaining blocks are commented examples using placeholder content; remove the leading hash and space from each line to enable.");
        builder.AppendLine("# Validate with: docxedit check input.docx template.docxpatch");
        builder.AppendLine("# Apply with: docxedit apply input.docx template.docxpatch --output out.docx");
        builder.Append("# Capabilities: docxedit capabilities input.docx --id ").Append(capabilities.TargetId)
            .Append(" --track-changes ").AppendLine(ModeWord(mode));
        builder.AppendLine("# Operation help: docxedit help <operation-name>");
        builder.AppendLine();
        builder.AppendLine("docxpatch 1");
        builder.AppendLine();
    }

    private static void AppendHeredoc(StringBuilder builder, string field, string value)
    {
        if (!DocxPatchWriter.IsSafeHeredocBody(value))
        {
            builder.AppendLine(DocxPatchWriter.WriteField(field, value));
            return;
        }

        builder.Append(field).AppendLine(" <<<");
        builder.AppendLine(value);
        builder.AppendLine(">>>");
    }

    private static void AppendBlock(StringBuilder builder, bool active, Action<StringBuilder> write)
    {
        if (active)
        {
            write(builder);
            builder.AppendLine();
            return;
        }

        var staged = new StringBuilder();
        write(staged);
        AppendCommented(builder, staged.ToString());
        builder.AppendLine();
    }

    private static void AppendCommented(StringBuilder builder, string text)
    {
        using var reader = new StringReader(text);
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            builder.Append("# ").AppendLine(line);
        }
    }

    private static void AppendConditionalNote(StringBuilder builder, DocxTargetCapabilities capabilities, string operation)
    {
        builder.Append("# Conditional: ").AppendLine(ReasonOf(capabilities, operation));
    }    private static string? BuildParagraphTemplate(StoryDocument storyDocument, DocxTargetId parsed, DocxTargetCapabilities capabilities, TrackChangesMode mode)
    {
        XElement? root = storyDocument.Document.Root;
        XElement container = root?.Element(OoxmlNs.W + "body") ?? root!;
        XElement? paragraph = DocxStoryBlocks.FindParagraphByPhysicalOrdinal(container, parsed.Primary);
        if (paragraph is null)
        {
            return null;
        }

        string target = parsed.ToWireValue();
        string current = ReadVisibleText(paragraph);
        string? currentStyle = (string?)paragraph
            .Element(OoxmlNs.W + "pPr")
            ?.Element(OoxmlNs.W + "pStyle")
            ?.Attribute(OoxmlNs.W + "val");
        string replaceSupport = SupportOf(capabilities, "replace-text");
        string replaceParagraphSupport = SupportOf(capabilities, "replace-paragraph");
        bool activeReplaceText = string.Equals(replaceSupport, "supported", StringComparison.Ordinal) && current.Length != 0;
        bool activeReplaceParagraph = !activeReplaceText && string.Equals(replaceParagraphSupport, "supported", StringComparison.Ordinal) && current.Length == 0;
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, activeReplaceText || activeReplaceParagraph);
        if (activeReplaceText)
        {
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op replace-text");
                block.AppendLine("target " + target);
                AppendHeredoc(block, "expect-text", current);
                AppendHeredoc(block, "find", current);
                AppendHeredoc(block, "with", current);
                block.AppendLine("end");
            });
        }
        else if (activeReplaceParagraph)
        {
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op replace-paragraph");
                block.AppendLine("target " + target);
                AppendHeredoc(block, "text", current);
                block.AppendLine("end");
            });
        }

        AppendOpBlock(builder, capabilities, "replace-text", replaceSupport, activeReplaceText, block =>
        {
            block.AppendLine("op replace-text");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "expect-text", current);
            AppendHeredoc(block, "find", current);
            AppendHeredoc(block, "with", "New text");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "replace-paragraph", replaceParagraphSupport, activeReplaceParagraph, block =>
        {
            block.AppendLine("op replace-paragraph");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "text", "New paragraph text");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-style", SupportOf(capabilities, "set-style"), active: false, block =>
        {
            block.AppendLine("op set-style");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "style", currentStyle ?? "Normal");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "insert-before", SupportOf(capabilities, "insert-before"), active: false, block =>
        {
            block.AppendLine("op insert-before");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "text", "Inserted paragraph");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "insert-after", SupportOf(capabilities, "insert-after"), active: false, block =>
        {
            block.AppendLine("op insert-after");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "text", "Inserted paragraph");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "delete-block", SupportOf(capabilities, "delete-block"), active: false, block =>
        {
            block.AppendLine("op delete-block");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "expect-text", current);
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "add-comment", SupportOf(capabilities, "add-comment"), active: false, block =>
        {
            block.AppendLine("op add-comment");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "text", "Review note");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "add-bookmark", SupportOf(capabilities, "add-bookmark"), active: false, block =>
        {
            block.AppendLine("op add-bookmark");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "name", "MarkName");
            block.AppendLine("end");
        });
        return builder.ToString();
    }

    // Emits nothing for unsupported operations; emits an active block when
    // requested and the verdict allows it; otherwise emits a commented block
    // with a condition note for conditional verdicts.
    private static void AppendOpBlock(StringBuilder builder, DocxTargetCapabilities capabilities, string operation, string support, bool active, Action<StringBuilder> write)
    {
        if (active)
        {
            return;
        }

        if (string.Equals(support, "unsupported", StringComparison.Ordinal))
        {
            return;
        }

        if (string.Equals(support, "conditional", StringComparison.Ordinal))
        {
            AppendConditionalNote(builder, capabilities, operation);
        }

        AppendBlock(builder, active: false, write);
    }
    private static string? BuildContentControlTemplate(StoryDocument storyDocument, DocxTargetId parsed, DocxTargetCapabilities capabilities, TrackChangesMode mode)
    {
        XElement? control = storyDocument.Document.Descendants(OoxmlNs.W + "sdt").ElementAtOrDefault(parsed.Primary - 1);
        XElement? content = control?.Element(OoxmlNs.W + "sdtContent");
        if (control is null || content is null)
        {
            return null;
        }

        string target = parsed.ToWireValue();
        string current = ReadVisibleText(content);
        bool isRichText = IsRichTextContentControl(control);
        string textSupport = SupportOf(capabilities, "set-content-control-text");
        bool activeText = string.Equals(textSupport, "supported", StringComparison.Ordinal);
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, activeText);
        if (activeText)
        {
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op set-content-control-text");
                block.AppendLine("target " + target);
                if (isRichText)
                {
                    AppendHeredoc(block, "expect-text", current);
                }

                AppendHeredoc(block, "text", current);
                block.AppendLine("end");
            });
        }

        AppendOpBlock(builder, capabilities, "set-content-control-text", textSupport, activeText, block =>
        {
            block.AppendLine("op set-content-control-text");
            block.AppendLine("target " + target);
            if (isRichText)
            {
                AppendHeredoc(block, "expect-text", current);
            }

            AppendHeredoc(block, "text", "New control text");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-content-control-checkbox", SupportOf(capabilities, "set-content-control-checkbox"), active: false, block =>
        {
            block.AppendLine("op set-content-control-checkbox");
            block.AppendLine("target " + target);
            block.AppendLine("checked true");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-content-control-choice", SupportOf(capabilities, "set-content-control-choice"), active: false, block =>
        {
            block.AppendLine("op set-content-control-choice");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "value", ReadFirstChoiceItemValue(control));
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-content-control-date", SupportOf(capabilities, "set-content-control-date"), active: false, block =>
        {
            block.AppendLine("op set-content-control-date");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "value", ReadControlDateValue(control));
            block.AppendLine("end");
        });
        return builder.ToString();
    }

    private static string ReadFirstChoiceItemValue(XElement control)
    {
        XElement? list = control
            .Element(OoxmlNs.W + "sdtPr")
            ?.Elements()
            .FirstOrDefault(element => element.Name == OoxmlNs.W + "dropDownList" || element.Name == OoxmlNs.W + "comboBox");
        XElement? item = list?.Elements(OoxmlNs.W + "listItem").FirstOrDefault();
        return (string?)item?.Attribute(OoxmlNs.W + "val")
            ?? (string?)item?.Attribute(OoxmlNs.W + "displayText")
            ?? "ITEM";
    }

    private static string ReadControlDateValue(XElement control)
    {
        return (string?)control
            .Element(OoxmlNs.W + "sdtPr")
            ?.Element(OoxmlNs.W + "date")
            ?.Element(OoxmlNs.W + "fullDate")
            ?.Attribute(OoxmlNs.W + "val")
            ?? "2026-01-01T00:00:00Z";
    }

    private static string? BuildBookmarkTemplate(StoryDocument storyDocument, DocxTargetId parsed, DocxTargetCapabilities capabilities, TrackChangesMode mode)
    {
        XElement? start = storyDocument.Document.Descendants(OoxmlNs.W + "bookmarkStart").ElementAtOrDefault(parsed.Primary - 1);
        string? ooxmlId = (string?)start?.Attribute(OoxmlNs.W + "id");
        XElement? end = ooxmlId is null
            ? null
            : storyDocument.Document.Descendants(OoxmlNs.W + "bookmarkEnd").FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), ooxmlId, StringComparison.Ordinal));
        bool sameParagraph = start is not null && end is not null &&
            start.Parent is not null &&
            start.Parent.Name == OoxmlNs.W + "p" &&
            start.Parent == end.Parent;
        if (start is null)
        {
            return null;
        }
        string target = parsed.ToWireValue();
        string current = sameParagraph
            ? ReadVisibleText(new XElement(OoxmlNs.W + "p", start.NodesAfterSelf().TakeWhile(node => node != end).ToArray()))
            : string.Empty;
        string name = (string?)start.Attribute(OoxmlNs.W + "name") ?? "MarkName";
        string replaceSupport = SupportOf(capabilities, "replace-bookmark-text");
        bool activeReplace = string.Equals(replaceSupport, "supported", StringComparison.Ordinal) && current.Length != 0;
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, activeReplace);
        if (activeReplace)
        {
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op replace-bookmark-text");
                block.AppendLine("target " + target);
                AppendHeredoc(block, "expect-text", current);
                AppendHeredoc(block, "text", current);
                block.AppendLine("end");
            });
        }
        AppendOpBlock(builder, capabilities, "replace-bookmark-text", replaceSupport, activeReplace, block =>
        {
            block.AppendLine("op replace-bookmark-text");
            block.AppendLine("target " + target);
            if (sameParagraph)
            {
                AppendHeredoc(block, "expect-text", current);
            }
            AppendHeredoc(block, "text", "New bookmark text");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "rename-bookmark", SupportOf(capabilities, "rename-bookmark"), active: false, block =>
        {
            block.AppendLine("op rename-bookmark");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "name", name);
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "delete-bookmark", SupportOf(capabilities, "delete-bookmark"), active: false, block =>
        {
            block.AppendLine("op delete-bookmark");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "expect-name", name);
            block.AppendLine("end");
        });
        return builder.ToString();
    }
    private static string? BuildTableTemplate(StoryDocument storyDocument, DocxTargetId parsed, DocxTargetCapabilities capabilities, TrackChangesMode mode)
    {
        XElement? root = storyDocument.Document.Root;
        XElement container = root?.Element(OoxmlNs.W + "body") ?? root!;
        XElement? table = DocxStoryBlocks.FindTableByPhysicalOrdinal(container, parsed.Primary);
        if (table is null)
        {
            return null;
        }
        string target = parsed.ToWireValue();
        string? caption = ReadTableTextProperty(table, "tblCaption");
        string? description = ReadTableTextProperty(table, "tblDescription");
        string? style = ReadTableStyleId(table);
        string current = ReadVisibleText(table);
        string metadataSupport = SupportOf(capabilities, "set-table-metadata");
        bool activeMetadata = string.Equals(metadataSupport, "supported", StringComparison.Ordinal) && (caption is not null || description is not null);
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, activeMetadata);
        if (activeMetadata)
        {
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op set-table-metadata");
                block.AppendLine("target " + target);
                if (caption is not null)
                {
                    AppendHeredoc(block, "caption", caption);
                }
                if (description is not null)
                {
                    AppendHeredoc(block, "description", description);
                }
                block.AppendLine("end");
            });
        }
        AppendOpBlock(builder, capabilities, "set-table-metadata", metadataSupport, activeMetadata, block =>
        {
            block.AppendLine("op set-table-metadata");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "caption", caption ?? "Table caption");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-table-style", SupportOf(capabilities, "set-table-style"), active: false, block =>
        {
            block.AppendLine("op set-table-style");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "style", style ?? "TableGrid");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "append-row", SupportOf(capabilities, "append-row"), active: false, block =>
        {
            block.AppendLine("op append-row");
            block.AppendLine("target " + target);
            XElement? templateRow = table.Elements(OoxmlNs.W + "tr").LastOrDefault();
            int cellCount = templateRow is null ? 0 : templateRow.Elements(OoxmlNs.W + "tc").Count();
            for (int index = 0; index < cellCount; index++)
            {
                AppendHeredoc(block, "cell", "New cell");
            }
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "insert-after", SupportOf(capabilities, "insert-after"), active: false, block =>
        {
            block.AppendLine("op insert-after");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "text", "Inserted paragraph");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "delete-block", SupportOf(capabilities, "delete-block"), active: false, block =>
        {
            block.AppendLine("op delete-block");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "expect-text", current);
            block.AppendLine("end");
        });
        return builder.ToString();
    }
    private static string? BuildRowTemplate(StoryDocument storyDocument, DocxTargetId parsed, DocxTargetCapabilities capabilities, TrackChangesMode mode)
    {
        XElement? root = storyDocument.Document.Root;
        XElement container = root?.Element(OoxmlNs.W + "body") ?? root!;
        XElement? table = DocxStoryBlocks.FindTableByPhysicalOrdinal(container, parsed.Primary);
        XElement? row = table is null || parsed.Secondary < 1
            ? null
            : table.Elements(OoxmlNs.W + "tr").ElementAtOrDefault(parsed.Secondary - 1);
        if (table is null || row is null)
        {
            return null;
        }
        string target = parsed.ToWireValue();
        bool currentHeader = ReadTableRowHeader(row);
        string current = ReadVisibleText(row);
        int cellCount = row.Elements(OoxmlNs.W + "tc").Count();
        string headerSupport = SupportOf(capabilities, "set-row-header");
        bool activeHeader = string.Equals(headerSupport, "supported", StringComparison.Ordinal);
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, activeHeader);
        if (activeHeader)
        {
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op set-row-header");
                block.AppendLine("target " + target);
                block.AppendLine("header " + (currentHeader ? "true" : "false"));
                block.AppendLine("end");
            });
        }
        AppendOpBlock(builder, capabilities, "insert-row-before", SupportOf(capabilities, "insert-row-before"), active: false, block =>
        {
            block.AppendLine("op insert-row-before");
            block.AppendLine("target " + target);
            for (int index = 0; index < cellCount; index++)
            {
                AppendHeredoc(block, "cell", "New cell");
            }
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "insert-row-after", SupportOf(capabilities, "insert-row-after"), active: false, block =>
        {
            block.AppendLine("op insert-row-after");
            block.AppendLine("target " + target);
            for (int index = 0; index < cellCount; index++)
            {
                AppendHeredoc(block, "cell", "New cell");
            }
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "delete-row", SupportOf(capabilities, "delete-row"), active: false, block =>
        {
            block.AppendLine("op delete-row");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "expect-contains", current);
            block.AppendLine("end");
        });
        return builder.ToString();
    }
    private static string? BuildSectionTemplate(StoryDocument storyDocument, DocxTargetId parsed, DocxTargetCapabilities capabilities, TrackChangesMode mode)
    {
        if (parsed.Story != (char)77 || parsed.Primary < 1)
        {
            return null;
        }
        XElement? sectionProperties = null;
        int sectionIndex = 0;
        foreach (DocxStoryBlocks.StoryBlock entry in DocxStoryBlocks.EnumeratePhysicalBlocks(storyDocument.Document.Root?.Element(OoxmlNs.W + "body") ?? storyDocument.Document.Root!))
        {
            XElement block = entry.Block;
            if (block.Name == OoxmlNs.W + "p")
            {
                XElement? paragraphSectPr = block.Element(OoxmlNs.W + "pPr")?.Element(OoxmlNs.W + "sectPr");
                if (paragraphSectPr is not null)
                {
                    sectionIndex++;
                    if (sectionIndex == parsed.Primary)
                    {
                        sectionProperties = paragraphSectPr;
                        break;
                    }
                }
            }
            else if (block.Name == OoxmlNs.W + "sectPr")
            {
                sectionIndex++;
                if (sectionIndex == parsed.Primary)
                {
                    sectionProperties = block;
                    break;
                }
            }
        }
        if (sectionProperties is null)
        {
            return null;
        }
        string target = parsed.ToWireValue();
        int currentColumns = ReadSectionColumnCount(sectionProperties);
        string currentOrientation = ReadSectionOrientation(sectionProperties).ToWireValue();
        bool hasTrackedRevision = sectionProperties.Elements(OoxmlNs.W + "sectPrChange").Any();
        string columnsSupport = SupportOf(capabilities, "set-section-columns");
        bool activeColumns = !hasTrackedRevision && string.Equals(columnsSupport, "conditional", StringComparison.Ordinal);
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, activeColumns);
        if (activeColumns)
        {
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op set-section-columns");
                block.AppendLine("target " + target);
                block.AppendLine("count " + currentColumns);
                block.AppendLine("end");
            });
        }
        AppendOpBlock(builder, capabilities, "set-section-columns", columnsSupport, activeColumns, block =>
        {
            block.AppendLine("op set-section-columns");
            block.AppendLine("target " + target);
            block.AppendLine("count " + currentColumns);
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-section-orientation", SupportOf(capabilities, "set-section-orientation"), active: false, block =>
        {
            block.AppendLine("op set-section-orientation");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "orientation", currentOrientation);
            block.AppendLine("end");
        });
        return builder.ToString();
    }
    private static string? BuildHyperlinkTemplate(StoryDocument storyDocument, DocxTargetId parsed, DocxTargetCapabilities capabilities, TrackChangesMode mode)
    {
        XElement? hyperlink = storyDocument.Document.Descendants(OoxmlNs.W + "hyperlink").ElementAtOrDefault(parsed.Primary - 1);
        if (hyperlink is null || parsed.Primary < 1)
        {
            return null;
        }
        string target = parsed.ToWireValue();
        string current = ReadVisibleText(hyperlink);
        string textSupport = SupportOf(capabilities, "set-hyperlink-text");
        bool activeText = !string.Equals(textSupport, "unsupported", StringComparison.Ordinal) && current.Length != 0;
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, activeText);
        if (activeText)
        {
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op set-hyperlink-text");
                block.AppendLine("target " + target);
                AppendHeredoc(block, "expect-text", current);
                AppendHeredoc(block, "text", current);
                block.AppendLine("end");
            });
        }
        AppendOpBlock(builder, capabilities, "set-hyperlink-text", textSupport, activeText, block =>
        {
            block.AppendLine("op set-hyperlink-text");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "expect-text", current);
            AppendHeredoc(block, "text", "New link text");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-hyperlink-target", SupportOf(capabilities, "set-hyperlink-target"), active: false, block =>
        {
            block.AppendLine("op set-hyperlink-target");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "uri", "https://example.test/new");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "remove-hyperlink", SupportOf(capabilities, "remove-hyperlink"), active: false, block =>
        {
            block.AppendLine("op remove-hyperlink");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "expect-text", current);
            block.AppendLine("end");
        });
        return builder.ToString();
    }
    private static string? BuildFieldTemplate(StoryDocument storyDocument, DocxTargetId parsed, DocxTargetCapabilities capabilities, TrackChangesMode mode)
    {
        XElement? root = storyDocument.Document.Root;
        if (root is null || parsed.Primary < 1)
        {
            return null;
        }
        XElement scope = root.Element(OoxmlNs.W + "body") ?? root;
        XElement? element = FindFields(scope).ElementAtOrDefault(parsed.Primary - 1);
        if (element is null)
        {
            return null;
        }
        string target = parsed.ToWireValue();
        bool simple = element.Name == OoxmlNs.W + "fldSimple";
        string current = string.Empty;
        if (simple)
        {
            current = ReadVisibleText(element);
        }
        else if (TryGetSimpleComplexFieldResultRuns(element, out _, out XElement[] resultRuns, out _))
        {
            current = ReadVisibleText(new XElement(OoxmlNs.W + "p", resultRuns));
        }
        string code = simple
            ? NormalizeFieldCodeForGuard((string?)element.Attribute(OoxmlNs.W + "instr") ?? string.Empty)
            : "New code";
        string resultSupport = SupportOf(capabilities, "set-field-result");
        bool activeResult = string.Equals(resultSupport, "supported", StringComparison.Ordinal) && current.Length != 0;
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, activeResult);
        if (activeResult)
        {
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op set-field-result");
                block.AppendLine("target " + target);
                AppendHeredoc(block, "expect-result", current);
                AppendHeredoc(block, "text", current);
                block.AppendLine("end");
            });
        }
        AppendOpBlock(builder, capabilities, "set-field-result", resultSupport, activeResult, block =>
        {
            block.AppendLine("op set-field-result");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "expect-result", current);
            AppendHeredoc(block, "text", "New result text");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-field-code", SupportOf(capabilities, "set-field-code"), active: false, block =>
        {
            block.AppendLine("op set-field-code");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "code", code);
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-field-dirty", SupportOf(capabilities, "set-field-dirty"), active: false, block =>
        {
            block.AppendLine("op set-field-dirty");
            block.AppendLine("target " + target);
            block.AppendLine("dirty true");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-field-lock", SupportOf(capabilities, "set-field-lock"), active: false, block =>
        {
            block.AppendLine("op set-field-lock");
            block.AppendLine("target " + target);
            block.AppendLine("locked true");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "refresh-field-result", SupportOf(capabilities, "refresh-field-result"), active: false, block =>
        {
            block.AppendLine("op refresh-field-result");
            block.AppendLine("target " + target);
            block.AppendLine("end");
        });
        return builder.ToString();
    }
    private static string? BuildImageTemplate(
        StoryDocument storyDocument,
        DocxTargetId parsed,
        DocxTargetCapabilities capabilities,
        TrackChangesMode mode,
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        ImageBlipTarget? imageTarget = FindImageBlipTarget(package, storyDocument.PartName, parsed.Story, parsed.StoryPart, parsed.Primary, cancellationToken, allowLiveFallback: true);
        if (imageTarget is null)
        {
            return null;
        }
        string target = parsed.ToWireValue();
        XElement? drawing = imageTarget.Blip.Ancestors(OoxmlNs.W + "drawing").FirstOrDefault();
        XElement? container = drawing?.Descendants(OoxmlNs.Wp + "inline").FirstOrDefault()
            ?? drawing?.Descendants(OoxmlNs.Wp + "anchor").FirstOrDefault();
        string? alt = (string?)container?.Element(OoxmlNs.Wp + "docPr")?.Attribute("descr");
        string altSupport = SupportOf(capabilities, "set-image-alt");
        bool activeAlt = string.Equals(altSupport, "supported", StringComparison.Ordinal) && alt is not null;
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, activeAlt);
        if (activeAlt)
        {
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op set-image-alt");
                block.AppendLine("target " + target);
                AppendHeredoc(block, "alt", alt!);
                block.AppendLine("end");
            });
        }
        AppendOpBlock(builder, capabilities, "set-image-alt", altSupport, activeAlt, block =>
        {
            block.AppendLine("op set-image-alt");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "alt", alt ?? "New alt text");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "replace-image", SupportOf(capabilities, "replace-image"), active: false, block =>
        {
            block.AppendLine("op replace-image");
            block.AppendLine("target " + target);
            AppendHeredoc(block, "asset", "new-image.png");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-image-size", SupportOf(capabilities, "set-image-size"), active: false, block =>
        {
            block.AppendLine("op set-image-size");
            block.AppendLine("target " + target);
            block.AppendLine("width 5cm");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "set-image-crop", SupportOf(capabilities, "set-image-crop"), active: false, block =>
        {
            block.AppendLine("op set-image-crop");
            block.AppendLine("target " + target);
            block.AppendLine("left-percent 10");
            block.AppendLine("end");
        });
        AppendOpBlock(builder, capabilities, "delete-image", SupportOf(capabilities, "delete-image"), active: false, block =>
        {
            block.AppendLine("op delete-image");
            block.AppendLine("target " + target);
            block.AppendLine("end");
        });
        return builder.ToString();
    }
    private static string? BuildCellTemplate(StoryDocument storyDocument, DocxTargetId parsed, DocxTargetCapabilities capabilities, TrackChangesMode mode, bool isMergeGroup)
    {
        XElement? root = storyDocument.Document.Root;
        XElement container = root?.Element(OoxmlNs.W + "body") ?? root!;
        XElement? table = DocxStoryBlocks.FindTableByPhysicalOrdinal(container, parsed.Primary);
        XElement? row = null;
        XElement? cell = null;
        if (table is null)
        {
            return null;
        }

        if (isMergeGroup)
        {
            if (!TryFindMergeGroupRoot(table, parsed.Secondary, out XElement? groupRow, out XElement? groupCell, out _, out _))
            {
                return null;
            }

            row = groupRow;
            cell = groupCell;
        }
        else
        {
            row = table.Elements(OoxmlNs.W + "tr").ElementAtOrDefault(parsed.Secondary - 1);
            cell = row is null ? null : FindCellByVisualColumn(row, parsed.Tertiary);
        }

        if (row is null || cell is null)
        {
            return null;
        }

        string target = parsed.ToWireValue();
        string current = ReadVisibleText(cell);
        bool simple = IsSimpleEditableCell(cell);
        string setSupport = SupportOf(capabilities, "set-cell");
        bool activeSet = string.Equals(setSupport, "supported", StringComparison.Ordinal)
            || string.Equals(setSupport, "conditional", StringComparison.Ordinal);
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, activeSet);
        if (activeSet)
        {
            bool force = !simple;
            AppendBlock(builder, active: true, block =>
            {
                block.AppendLine("op set-cell");
                block.AppendLine("target " + target);
                AppendHeredoc(block, "expect-text", current);
                AppendHeredoc(block, "text", current);
                if (force)
                {
                    block.AppendLine("force true");
                }

                block.AppendLine("end");
            });
        }

        if (!activeSet)
        {
            AppendOpBlock(builder, capabilities, "set-cell", setSupport, active: false, block =>
            {
                block.AppendLine("op set-cell");
                block.AppendLine("target " + target);
                AppendHeredoc(block, "expect-text", current);
                AppendHeredoc(block, "text", "New cell text");
                block.AppendLine("end");
            });
        }

        AppendOpBlock(builder, capabilities, "set-cell-shading", SupportOf(capabilities, "set-cell-shading"), active: false, block =>
        {
            block.AppendLine("op set-cell-shading");
            block.AppendLine("target " + target);
            block.AppendLine("fill 4472C4");
            block.AppendLine("end");
        });
        return builder.ToString();
    }
}
