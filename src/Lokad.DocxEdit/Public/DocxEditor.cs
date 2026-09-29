using System.Diagnostics.CodeAnalysis;
using Lokad.DocxEdit.Ooxml;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Rendering;
using System.Xml;
using System.Text;

namespace Lokad.DocxEdit;

/// <summary>
/// Stream-first entry point for inspecting and editing .docx packages.
/// </summary>
public sealed class DocxEditor
{
    /// <summary>Reads a document. Uses default options and no cancellation.</summary>
    public DocxReadResult Read(
        Stream input)
    {
        return Read(input, new DocxReadOptions(), CancellationToken.None);
    }

    /// <summary>Reads a document. Uses no cancellation.</summary>
    public DocxReadResult Read(
        Stream input, DocxReadOptions options)
    {
        return Read(input, options, CancellationToken.None);
    }

    /// <summary>Reads a document with explicit options and cancellation.</summary>
    public DocxReadResult Read(
        Stream input,
        DocxReadOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxReadResult
            {
                Success = false,
                Diagnostics = diagnostics
            };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, options.IncludeHeadersFooters, options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxReadResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                PartNames = SortedPartNames(package),
                MainDocumentPartName = package.MainDocumentPartName
            };
        }

        DocxDocumentModel redacted = ApplyMaxTextToReadModel(model, options.MaxText);
        if (!TryUnsupportedScan(package, options.IncludeHeadersFooters, cancellationToken, out IReadOnlyList<DocxDiagnostic>? readUnsupported, out IReadOnlyList<DocxDiagnostic> readUnsupportedFailure))
        {
            return new DocxReadResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(readUnsupportedFailure).ToArray(),
                PartNames = SortedPartNames(package),
                MainDocumentPartName = package.MainDocumentPartName
            };
        }

        return new DocxReadResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(readUnsupported)
                .ToArray(),
            PartNames = SortedPartNames(package),
            MainDocumentPartName = package.MainDocumentPartName,
            Text = TextRenderers.RenderRead(redacted, options.MaxText),
            Paragraphs = redacted.Paragraphs,
            Tables = redacted.Tables,
            Images = redacted.Images,
            Sections = redacted.Sections,
            Bookmarks = redacted.Bookmarks,
            ContentControls = redacted.ContentControls,
            Fields = redacted.Fields,
            Hyperlinks = redacted.Hyperlinks
        };
    }

    /// <summary>
    /// Applies <c>MaxText</c> to per-item body text in a scanned model.
    /// Contract: paragraph/run/cell text, field cached results, and table/image
    /// captions/descriptions/titles are truncated (0 drops); IDs, counts, style
    /// IDs/names, bookmark names, content-control tags/aliases, field codes/kinds,
    /// hyperlink URIs/anchors, authors, and revision IDs are retained as
    /// structural metadata. Length/count properties keep full-text values.
    /// </summary>
    private static DocxDocumentModel ApplyMaxTextToReadModel(DocxDocumentModel model, int maxText)
    {
        IReadOnlyList<DocxParagraphInfo> paragraphs = model.Paragraphs
            .Select(paragraph => paragraph with
            {
                Text = TextRenderers.Truncate(paragraph.Text, maxText),
                Runs = paragraph.Runs
                    .Select(run => run with { Text = TextRenderers.Truncate(run.Text, maxText) })
                    .ToArray()
            })
            .ToArray();
        IReadOnlyList<DocxTableInfo> tables = model.Tables
            .Select(table => table with
            {
                Caption = table.Caption is null ? null : TextRenderers.Truncate(table.Caption, maxText),
                Description = table.Description is null ? null : TextRenderers.Truncate(table.Description, maxText),
                Cells = table.Cells
                    .Select(cell => cell with { Text = TextRenderers.Truncate(cell.Text, maxText) })
                    .ToArray()
            })
            .ToArray();
        IReadOnlyList<DocxFieldInfo> fields = model.Fields
            .Select(field => field with { CachedResultText = TextRenderers.Truncate(field.CachedResultText, maxText) })
            .ToArray();
        IReadOnlyList<DocxImageInfo> images = model.Images
            .Select(image => image with
            {
                Description = image.Description is null ? null : TextRenderers.Truncate(image.Description, maxText),
                Title = image.Title is null ? null : TextRenderers.Truncate(image.Title, maxText)
            })
            .ToArray();
        return new DocxDocumentModel(paragraphs, tables, images, model.Sections, model.Bookmarks, model.ContentControls, fields, model.Hyperlinks);
    }

    /// <summary>Outlines a document. Uses default options and no cancellation.</summary>
    public DocxOutlineResult Outline(
        Stream input)
    {
        return Outline(input, new DocxOutlineOptions(), CancellationToken.None);
    }

    /// <summary>Outlines a document. Uses no cancellation.</summary>
    public DocxOutlineResult Outline(
        Stream input, DocxOutlineOptions options)
    {
        return Outline(input, options, CancellationToken.None);
    }

    /// <summary>Outlines document structure with explicit options and cancellation.</summary>
    public DocxOutlineResult Outline(
        Stream input,
        DocxOutlineOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxOutlineResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, options.IncludeHeadersFooters, options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxOutlineResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                PartNames = SortedPartNames(package),
                MainDocumentPartName = package.MainDocumentPartName
            };
        }

        if (!TryUnsupportedScan(package, options.IncludeHeadersFooters, cancellationToken, out IReadOnlyList<DocxDiagnostic>? outlineUnsupported, out IReadOnlyList<DocxDiagnostic> outlineUnsupportedFailure))
        {
            return new DocxOutlineResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(outlineUnsupportedFailure).ToArray(),
                PartNames = SortedPartNames(package),
                MainDocumentPartName = package.MainDocumentPartName
            };
        }

        return new DocxOutlineResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(outlineUnsupported)
                .ToArray(),
            PartNames = SortedPartNames(package),
            MainDocumentPartName = package.MainDocumentPartName,
            Items = TextRenderers.BuildOutline(model, options.MaxText)
        };
    }

    /// <summary>Finds text in a document. Uses default options and no cancellation.</summary>
    public DocxFindResult Find(
        Stream input, string query)
    {
        return Find(input, query, new DocxFindOptions(), CancellationToken.None);
    }

    /// <summary>Finds text in a document. Uses no cancellation.</summary>
    public DocxFindResult Find(
        Stream input, string query, DocxFindOptions options)
    {
        return Find(input, query, options, CancellationToken.None);
    }

    /// <summary>Searches visible text with explicit options and cancellation.</summary>
    public DocxFindResult Find(
        Stream input,
        string query,
        DocxFindOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(query);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxFindResult { Success = false, Diagnostics = diagnostics, Query = query };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, options.IncludeHeadersFooters, options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxFindResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                Query = query
            };
        }

        if (!TryUnsupportedScan(package, options.IncludeHeadersFooters, cancellationToken, out IReadOnlyList<DocxDiagnostic>? findUnsupported, out IReadOnlyList<DocxDiagnostic> findUnsupportedFailure))
        {
            return new DocxFindResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(findUnsupportedFailure).ToArray(),
                Query = query
            };
        }

        return new DocxFindResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(findUnsupported)
                .ToArray(),
            Query = query,
            Matches = TextRenderers.Find(model, query, options.MaxText)
        };
    }

    /// <summary>Whether a dump/context target ID addresses a header or footer story, which requires scanning those parts.</summary>
    private static bool TargetRequestsHeadersFooters(string targetId)
    {
        return DocxTargetId.TryParse(targetId, out DocxTargetId parsed) && parsed.Story != 'M';
    }

    /// <summary>Dumps one target. Uses default options and no cancellation.</summary>
    public DocxDumpResult Dump(
        Stream input, string targetId)
    {
        return Dump(input, targetId, new DocxDumpOptions(), CancellationToken.None);
    }

    /// <summary>Dumps one target. Uses no cancellation.</summary>
    public DocxDumpResult Dump(
        Stream input, string targetId, DocxDumpOptions options)
    {
        return Dump(input, targetId, options, CancellationToken.None);
    }

    /// <summary>Dumps one target with explicit options and cancellation.</summary>
    public DocxDumpResult Dump(
        Stream input,
        string targetId,
        DocxDumpOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetId);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxDumpResult { Success = false, Diagnostics = diagnostics, TargetId = targetId };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, includeHeadersFooters: TargetRequestsHeadersFooters(targetId), textView: options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxDumpResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                TargetId = targetId
            };
        }

        if (!TryDocumentOperation(() => DocxChangeScanner.Scan(package, includeCommentText: false, maxCommentText: 0, cancellationToken), out IReadOnlyList<DocxChangeInfo>? changes, out IReadOnlyList<DocxDiagnostic> changeDiagnostics))
        {
            return new DocxDumpResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(changeDiagnostics).ToArray(),
                TargetId = targetId
            };
        }

        if (!TryUnsupportedScan(package, includeHeadersFooters: TargetRequestsHeadersFooters(targetId), cancellationToken, out IReadOnlyList<DocxDiagnostic>? dumpUnsupported, out IReadOnlyList<DocxDiagnostic> dumpUnsupportedFailure))
        {
            return new DocxDumpResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(dumpUnsupportedFailure).ToArray(),
                TargetId = targetId
            };
        }

        string? text = TextRenderers.Dump(model, changes, targetId, options.IncludeRuns, options.MaxText);
        IReadOnlyList<DocxDumpRunInfo> runs = options.IncludeRuns ? TextRenderers.DumpRuns(model, targetId, options.MaxText) : [];
        if (text is null)
        {
            return new DocxDumpResult
            {
                Success = false,
                Diagnostics = diagnostics
                    .Concat(dumpUnsupported)
                    .Append(new DocxDiagnostic(DocxSeverity.Error, "E1201", TextRenderers.BuildUnknownTargetMessage(model, changes, targetId)) with { TargetId = targetId })
                    .ToArray(),
                TargetId = targetId
            };
        }

        return new DocxDumpResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(dumpUnsupported)
                .ToArray(),
            TargetId = targetId,
            Text = text,
            Runs = runs
        };
    }

    /// <summary>Summarizes the neighborhood of one target. Uses default options and no cancellation.</summary>
    public DocxContextResult Context(
        Stream input, string targetId)
    {
        return Context(input, targetId, new DocxContextOptions(), CancellationToken.None);
    }

    /// <summary>Summarizes the neighborhood of one target. Uses no cancellation.</summary>
    public DocxContextResult Context(
        Stream input, string targetId, DocxContextOptions options)
    {
        return Context(input, targetId, options, CancellationToken.None);
    }

    /// <summary>Summarizes target context with explicit options and cancellation.</summary>
    public DocxContextResult Context(
        Stream input,
        string targetId,
        DocxContextOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetId);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxContextResult { Success = false, Diagnostics = diagnostics, TargetId = targetId };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, (options.IncludeHeadersFooters || TargetRequestsHeadersFooters(targetId)), options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxContextResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                TargetId = targetId
            };
        }

        if (!TryDocumentOperation(() => DocxChangeScanner.Scan(package, includeCommentText: false, maxCommentText: 0, cancellationToken), out IReadOnlyList<DocxChangeInfo>? changes, out IReadOnlyList<DocxDiagnostic> changeDiagnostics))
        {
            return new DocxContextResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(changeDiagnostics).ToArray(),
                TargetId = targetId
            };
        }

        IReadOnlyList<DocxContextItem> items = TextRenderers.Context(model, changes, targetId, options.Radius, options.MaxText);
        IReadOnlyList<DocxDiagnostic> contextDiagnostics = items.Count == 0
            ? [new DocxDiagnostic(DocxSeverity.Error, "E1201", TextRenderers.BuildUnknownTargetMessage(model, changes, targetId)) with { TargetId = targetId }]
            : [];
        if (!TryUnsupportedScan(package, (options.IncludeHeadersFooters || TargetRequestsHeadersFooters(targetId)), cancellationToken, out IReadOnlyList<DocxDiagnostic>? contextUnsupported, out IReadOnlyList<DocxDiagnostic> contextUnsupportedFailure))
        {
            return new DocxContextResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(contextUnsupportedFailure).ToArray(),
                TargetId = targetId
            };
        }

        return new DocxContextResult
        {
            Success = items.Count > 0,
            Diagnostics = diagnostics
                .Concat(contextUnsupported)
                .Concat(contextDiagnostics)
                .ToArray(),
            TargetId = targetId,
            Items = items,
            Text = TextRenderers.RenderContext(items)
        };
    }

    /// <summary>Describes editing capabilities for one target (paragraph, content-control, cell, merge-group, bookmark, table, row, section, hyperlink, field, or image). Uses default options and no cancellation.</summary>
    public DocxCapabilitiesResult GetCapabilities(
        Stream input, string targetId)
    {
        return GetCapabilities(input, targetId, new DocxCapabilitiesOptions(), CancellationToken.None);
    }

    /// <summary>Describes editing capabilities for one target (paragraph, content-control, cell, merge-group, bookmark, table, row, section, hyperlink, field, or image). Uses no cancellation.</summary>
    public DocxCapabilitiesResult GetCapabilities(
        Stream input, string targetId, DocxCapabilitiesOptions options)
    {
        return GetCapabilities(input, targetId, options, CancellationToken.None);
    }

    /// <summary>Describes editing capabilities for one target (paragraph, content-control, cell, merge-group, bookmark, table, row, section, hyperlink, field, or image) with explicit options and cancellation.</summary>
    public DocxCapabilitiesResult GetCapabilities(
        Stream input,
        string targetId,
        DocxCapabilitiesOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetId);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxCapabilitiesResult { Success = false, Diagnostics = diagnostics, TargetId = targetId };
        }

        if (!TryUnsupportedScan(package, TargetRequestsHeadersFooters(targetId), cancellationToken, out IReadOnlyList<DocxDiagnostic>? capabilitiesUnsupported, out IReadOnlyList<DocxDiagnostic> capabilitiesUnsupportedFailure))
        {
            return new DocxCapabilitiesResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(capabilitiesUnsupportedFailure).ToArray(),
                TargetId = targetId
            };
        }

        if (!TryDocumentOperation(() => DocxPatchEngine.GetParagraphCapabilities(package, targetId, options.TrackChanges, cancellationToken), out DocxPatchEngine.ParagraphCapabilitiesOutcome? outcome, out IReadOnlyList<DocxDiagnostic> capabilitiesFailure))
        {
            return new DocxCapabilitiesResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(capabilitiesFailure).ToArray(),
                TargetId = targetId
            };
        }

        if (outcome is null || outcome.Capabilities is null)
        {
            DocxDiagnostic missing = outcome is null || outcome.Error is null
                ? new DocxDiagnostic(DocxSeverity.Error, "E1201", "Target " + ((char)39).ToString() + targetId + ((char)39).ToString() + " was not found.") with { TargetId = targetId }
                : outcome.Error with { TargetId = targetId };
            return new DocxCapabilitiesResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(capabilitiesUnsupported).Append(missing).ToArray(),
                TargetId = targetId
            };
        }

        return new DocxCapabilitiesResult
        {
            Success = true,
            Diagnostics = diagnostics.Concat(capabilitiesUnsupported).ToArray(),
            TargetId = targetId,
            Capabilities = outcome.Capabilities
        };
    }

    /// <summary>Builds a guarded patch template for one target. Uses default options and no cancellation.</summary>
    public DocxTemplateResult GetTemplate(
        Stream input, string targetId)
    {
        return GetTemplate(input, targetId, new DocxTemplateOptions(), CancellationToken.None);
    }

    /// <summary>Builds a guarded patch template for one target. Uses no cancellation.</summary>
    public DocxTemplateResult GetTemplate(
        Stream input, string targetId, DocxTemplateOptions options)
    {
        return GetTemplate(input, targetId, options, CancellationToken.None);
    }

    /// <summary>Builds a guarded patch template for one target with explicit options and cancellation.</summary>
    public DocxTemplateResult GetTemplate(
        Stream input,
        string targetId,
        DocxTemplateOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetId);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxTemplateResult { Success = false, Diagnostics = diagnostics, TargetId = targetId };
        }

        if (!TryUnsupportedScan(package, TargetRequestsHeadersFooters(targetId), cancellationToken, out IReadOnlyList<DocxDiagnostic>? templateUnsupported, out IReadOnlyList<DocxDiagnostic> templateUnsupportedFailure))
        {
            return new DocxTemplateResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(templateUnsupportedFailure).ToArray(),
                TargetId = targetId
            };
        }

        if (!TryDocumentOperation(() => DocxPatchEngine.GetTargetTemplate(package, targetId, options.TrackChanges, cancellationToken), out DocxPatchEngine.TargetTemplateOutcome? outcome, out IReadOnlyList<DocxDiagnostic> templateFailure))
        {
            return new DocxTemplateResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(templateFailure).ToArray(),
                TargetId = targetId
            };
        }

        if (outcome is null || outcome.Template is null)
        {
            DocxDiagnostic missing = outcome is null || outcome.Error is null
                ? new DocxDiagnostic(DocxSeverity.Error, "E1201", "Target " + ((char)39).ToString() + targetId + ((char)39).ToString() + " was not found.") with { TargetId = targetId }
                : outcome.Error with { TargetId = targetId };
            return new DocxTemplateResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(templateUnsupported).Append(missing).ToArray(),
                TargetId = targetId
            };
        }

        return new DocxTemplateResult
        {
            Success = true,
            Diagnostics = diagnostics.Concat(templateUnsupported).ToArray(),
            TargetId = targetId,
            Template = outcome.Template,
            Capabilities = outcome.Capabilities
        };
    }

    /// <summary>Lists styles. Uses default options and no cancellation.</summary>
    public DocxStylesResult Styles(
        Stream input)
    {
        return Styles(input, new DocxStylesOptions(), CancellationToken.None);
    }

    /// <summary>Lists styles. Uses no cancellation.</summary>
    public DocxStylesResult Styles(
        Stream input, DocxStylesOptions options)
    {
        return Styles(input, options, CancellationToken.None);
    }

    /// <summary>Lists styles with explicit options and cancellation.</summary>
    public DocxStylesResult Styles(
        Stream input,
        DocxStylesOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxStylesResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxStyleScanner.Scan(package, cancellationToken), out IReadOnlyList<DocxStyleInfo>? styles, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxStylesResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray()
            };
        }

        return new DocxStylesResult
        {
            Success = true,
            Diagnostics = diagnostics,
            Styles = styles
        };
    }

    /// <summary>Lists embedded images. Uses default options and no cancellation.</summary>
    public DocxMediaResult Media(
        Stream input)
    {
        return Media(input, new DocxMediaOptions(), CancellationToken.None);
    }

    /// <summary>Lists embedded images. Uses no cancellation.</summary>
    public DocxMediaResult Media(
        Stream input, DocxMediaOptions options)
    {
        return Media(input, options, CancellationToken.None);
    }

    /// <summary>Lists embedded images with explicit options and cancellation.</summary>
    public DocxMediaResult Media(
        Stream input,
        DocxMediaOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxMediaResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, includeHeadersFooters: false, textView: DocxTextView.Final, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxMediaResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray()
            };
        }

        if (!TryUnsupportedScan(package, includeHeadersFooters: false, cancellationToken, out IReadOnlyList<DocxDiagnostic>? mediaUnsupported, out IReadOnlyList<DocxDiagnostic> mediaUnsupportedFailure))
        {
            return new DocxMediaResult { Success = false, Diagnostics = diagnostics.Concat(mediaUnsupportedFailure).ToArray() };
        }

        return new DocxMediaResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(mediaUnsupported)
                .ToArray(),
            Images = model.Images
        };
    }

    /// <summary>Reads embedded image bytes. Uses default options and no cancellation.</summary>
    public DocxMediaExtractResult ExtractMedia(
        Stream input)
    {
        return ExtractMedia(input, new DocxMediaOptions(), CancellationToken.None);
    }

    /// <summary>Reads embedded image bytes. Uses no cancellation.</summary>
    public DocxMediaExtractResult ExtractMedia(
        Stream input, DocxMediaOptions options)
    {
        return ExtractMedia(input, options, CancellationToken.None);
    }

    /// <summary>Reads embedded image bytes with explicit options and cancellation.</summary>
    public DocxMediaExtractResult ExtractMedia(
        Stream input,
        DocxMediaOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxMediaExtractResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, includeHeadersFooters: false, textView: DocxTextView.Final, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxMediaExtractResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray()
            };
        }

        var files = new List<DocxMediaFile>();
        foreach (DocxImageInfo image in model.Images)
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? part = package.GetPart(image.PartName);
            if (part is null)
            {
                throw new InvalidDataException($"Inventoried image part '{image.PartName}' does not exist.");
            }

            files.Add(new DocxMediaFile(image.Id, image.PartName, $"{image.Id.ToWireValue()}-{Path.GetFileName(image.PartName)}", part.Bytes.ToArray()));
        }

        if (!TryUnsupportedScan(package, includeHeadersFooters: false, cancellationToken, out IReadOnlyList<DocxDiagnostic>? extractUnsupported, out IReadOnlyList<DocxDiagnostic> extractUnsupportedFailure))
        {
            return new DocxMediaExtractResult { Success = false, Diagnostics = diagnostics.Concat(extractUnsupportedFailure).ToArray() };
        }

        return new DocxMediaExtractResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(extractUnsupported)
                .ToArray(),
            Files = files
        };
    }

    /// <summary>Validates a document. Uses default options and no cancellation.</summary>
    public DocxValidateResult Validate(
        Stream input)
    {
        return Validate(input, new DocxValidateOptions(), CancellationToken.None);
    }

    /// <summary>Validates a document. Uses no cancellation.</summary>
    public DocxValidateResult Validate(
        Stream input, DocxValidateOptions options)
    {
        return Validate(input, options, CancellationToken.None);
    }

    /// <summary>Validates the package with explicit options and cancellation.</summary>
    public DocxValidateResult Validate(
        Stream input,
        DocxValidateOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxValidateResult { Success = false, Diagnostics = diagnostics, Profile = options.Profile };
        }

        if (!TryDocumentOperation(() => DocxPackageValidator.Validate(package, options.Profile, cancellationToken), out IReadOnlyList<DocxDiagnostic>? validationDiagnostics, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxValidateResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                Profile = options.Profile,
                PartNames = SortedPartNames(package),
                MainDocumentPartName = package.MainDocumentPartName
            };
        }

        IReadOnlyList<DocxDiagnostic> allDiagnostics = CapValidationDiagnostics(
            diagnostics.Concat(validationDiagnostics),
            options.MaxDiagnostics);
        return new DocxValidateResult
        {
            Success = allDiagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error),
            Diagnostics = allDiagnostics,
            Profile = options.Profile,
            PartNames = SortedPartNames(package),
            MainDocumentPartName = package.MainDocumentPartName
        };

        static IReadOnlyList<DocxDiagnostic> CapValidationDiagnostics(
            IEnumerable<DocxDiagnostic> diagnostics,
            int maxDiagnostics)
        {
            int cap = Math.Max(1, maxDiagnostics);
            DocxDiagnostic[] allDiagnostics = diagnostics.ToArray();
            if (allDiagnostics.Length <= cap)
            {
                return allDiagnostics;
            }

            DocxDiagnostic[] kept = allDiagnostics.Take(Math.Max(0, cap - 1)).ToArray();
            DocxDiagnostic[] omitted = allDiagnostics.Skip(kept.Length).ToArray();
            bool omittedErrors = omitted.Any(diagnostic => diagnostic.Severity == DocxSeverity.Error);
            DocxDiagnostic truncationDiagnostic = new(
                omittedErrors ? DocxSeverity.Error : DocxSeverity.Warning,
                omittedErrors ? "E9199" : "W9199",
                $"Validation diagnostics were capped at {cap}; omitted {omitted.Length} of {allDiagnostics.Length} diagnostic(s).");
            return kept.Append(truncationDiagnostic).ToArray();
        }
    }

    /// <summary>Lists tracked-change and comment markup. Uses default options and no cancellation.</summary>
    public DocxChangesResult Changes(
        Stream input)
    {
        return Changes(input, new DocxChangesOptions(), CancellationToken.None);
    }

    /// <summary>Lists tracked-change and comment markup. Uses no cancellation.</summary>
    public DocxChangesResult Changes(
        Stream input, DocxChangesOptions options)
    {
        return Changes(input, options, CancellationToken.None);
    }

    /// <summary>Lists existing markup with explicit options and cancellation.</summary>
    public DocxChangesResult Changes(
        Stream input,
        DocxChangesOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxChangesResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxChangeScanner.Scan(package, options.IncludeCommentText, options.MaxCommentText, cancellationToken), out IReadOnlyList<DocxChangeInfo>? changes, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxChangesResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                PartNames = SortedPartNames(package),
                MainDocumentPartName = package.MainDocumentPartName
            };
        }

        IReadOnlyList<DocxChangeInfo> annotatedChanges = AnnotateOperationReports(changes, options.OperationReports);

        return new DocxChangesResult
        {
            Success = true,
            Diagnostics = diagnostics,
            PartNames = SortedPartNames(package),
            MainDocumentPartName = package.MainDocumentPartName,
            Changes = annotatedChanges,
            Summary = DocxChangeScanner.Summarize(annotatedChanges),
            GroupSummary = DocxChangeScanner.SummarizeGroups(annotatedChanges),
            TargetSummary = DocxChangeScanner.SummarizeTargets(annotatedChanges),
            CommentSummary = DocxChangeScanner.SummarizeComments(annotatedChanges)
        };

        static IReadOnlyList<DocxChangeInfo> AnnotateOperationReports(
            IReadOnlyList<DocxChangeInfo> changes,
            IReadOnlyList<DocxPatchOperationReport> operationReports)
        {
            if (changes.Count == 0 || operationReports.Count == 0)
            {
                return changes;
            }

            var operationsByRevisionId = new Dictionary<string, DocxPatchOperationReport>(StringComparer.Ordinal);
            foreach (DocxPatchOperationReport operation in operationReports.Where(operation => operation.Success))
            {
                foreach (string revisionId in operation.GeneratedRevisionIds)
                {
                    if (!string.IsNullOrWhiteSpace(revisionId))
                    {
                        operationsByRevisionId.TryAdd(revisionId, operation);
                    }
                }
            }

            if (operationsByRevisionId.Count == 0)
            {
                return changes;
            }

            var annotated = new DocxChangeInfo[changes.Count];
            for (int i = 0; i < changes.Count; i++)
            {
                DocxChangeInfo change = changes[i];
                annotated[i] = change.RevisionId is not null &&
                    operationsByRevisionId.TryGetValue(change.RevisionId, out DocxPatchOperationReport? operation)
                        ? change with
                        {
                            OperationIndex = operation.Index,
                            OperationName = operation.OperationName,
                            OperationTarget = operation.Target
                        }
                        : change;
            }

            return annotated;
        }
    }

    /// <summary>Validates patch shape without loading a document: syntax, required fields, alternative groups, and exclusive fields. Uses no cancellation.</summary>
    public DocxLintResult Lint(
        TextReader patchReader)
    {
        return Lint(patchReader, CancellationToken.None);
    }

    /// <summary>Validates patch shape without loading a document with explicit cancellation.</summary>
    public DocxLintResult Lint(
        TextReader patchReader,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(patchReader);
        DocxPatch patch = ParsePatch(patchReader, new DocxEditOptions().MaxPatchChars, cancellationToken);
        if (!patch.Success)
        {
            return new DocxLintResult { Success = false, Diagnostics = patch.Diagnostics };
        }

        List<DocxDiagnostic> shape = DocxPatchEngine.ValidatePatchFields(patch).ToList();
        shape.AddRange(DocxPatchEngine.ValidatePatchAliases(patch));
        return new DocxLintResult
        {
            Success = shape.All(static diagnostic => diagnostic.Severity != DocxSeverity.Error),
            Diagnostics = shape,
            OperationCount = patch.Operations.Count,
            Operations = patch.Operations.Select(static operation => new DocxLintOperation(operation.Index, operation.OperationName)).ToArray(),
        };
    }

    /// <summary>Parses a patch. Uses no cancellation.</summary>
    public DocxPatch ParsePatch(
        TextReader patchReader)
    {
        return ParsePatch(patchReader, CancellationToken.None);
    }

    /// <summary>Parses patch text with explicit cancellation.</summary>
    public DocxPatch ParsePatch(
        TextReader patchReader,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(patchReader);
        return ParsePatch(patchReader, new DocxEditOptions().MaxPatchChars, cancellationToken);
    }

    private static DocxPatch ParsePatch(
        TextReader patchReader,
        int maxPatchChars,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var builder = new StringBuilder();
        var buffer = new char[8192];
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = patchReader.Read(buffer, 0, buffer.Length);
            if (read == 0)
            {
                break;
            }

            if ((long)builder.Length + read > maxPatchChars)
            {
                return new DocxPatch(false, 0, [], [new DocxDiagnostic(DocxSeverity.Error, "E2014", $"Patch exceeds the maximum supported size of {maxPatchChars} characters.")]);
            }

            builder.Append(buffer, 0, read);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return DocxPatchParser.Parse(builder.ToString());
    }

    /// <summary>Dry-runs a patch. Uses default options and no cancellation.</summary>
    public DocxCheckResult Check(
        Stream input, TextReader patchReader)
    {
        return Check(input, patchReader, new DocxEditOptions(), CancellationToken.None);
    }

    /// <summary>Dry-runs a patch. Uses no cancellation.</summary>
    public DocxCheckResult Check(
        Stream input, TextReader patchReader, DocxEditOptions options)
    {
        return Check(input, patchReader, options, CancellationToken.None);
    }

    /// <summary>Dry-runs a patch with explicit options and cancellation.</summary>
    public DocxCheckResult Check(
        Stream input,
        TextReader patchReader,
        DocxEditOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(patchReader);

        // The patch reader stays caller-owned. The input is disposed on every
        // exit below if and only if LeaveInputOpen is false (the loader also
        // disposes it on its own paths; disposal is idempotent).
        try
        {
            // Revision metadata is normalized once here; the engine and every result below share the effective values.
            DocxEditOptions effectiveOptions = options.WithNormalizedRevisionMetadata();
            DocxPatch patch = ParsePatch(patchReader, options.MaxPatchChars, cancellationToken);
            if (!patch.Success)
            {
                return new DocxCheckResult
                {
                    Success = false,
                    Diagnostics = patch.Diagnostics,
                    Author = effectiveOptions.Author,
                    TrackChanges = effectiveOptions.TrackChanges,
                    TimestampUtc = effectiveOptions.TimestampUtc
                };
            }

            OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: options.AllowMacroEnabledDocuments), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
            if (package is null)
            {
                return new DocxCheckResult
                {
                    Success = false,
                    Diagnostics = diagnostics,
                    Author = effectiveOptions.Author,
                    TrackChanges = effectiveOptions.TrackChanges,
                    TimestampUtc = effectiveOptions.TimestampUtc
                };
            }

            if (!TryDocumentOperation(() => DocxPatchEngine.Check(package, patch, effectiveOptions, cancellationToken), out PatchExecutionResult? execution, out IReadOnlyList<DocxDiagnostic> executionDiagnostics))
            {
                return new DocxCheckResult
                {
                    Success = false,
                    Diagnostics = diagnostics.Concat(executionDiagnostics).ToArray(),
                    Author = effectiveOptions.Author,
                    TrackChanges = effectiveOptions.TrackChanges,
                    TimestampUtc = effectiveOptions.TimestampUtc
                };
            }

            return new DocxCheckResult
            {
                Success = execution.Success,
                Diagnostics = diagnostics.Concat(execution.Diagnostics).ToArray(),
                Operations = execution.Reports,
                    Author = effectiveOptions.Author,
                    TrackChanges = effectiveOptions.TrackChanges,
                    TimestampUtc = effectiveOptions.TimestampUtc
            };
        }
        finally
        {
            if (!options.LeaveInputOpen)
            {
                input.Dispose();
            }
        }
    }

    /// <summary>Applies a patch. Uses default options and no cancellation.</summary>
    public DocxApplyResult Apply(
        Stream input, TextReader patchReader, Stream output)
    {
        return Apply(input, patchReader, output, new DocxEditOptions(), CancellationToken.None);
    }

    /// <summary>Applies a patch. Uses no cancellation.</summary>
    public DocxApplyResult Apply(
        Stream input, TextReader patchReader, Stream output, DocxEditOptions options)
    {
        return Apply(input, patchReader, output, options, CancellationToken.None);
    }

    /// <summary>Applies a patch with explicit options and cancellation.</summary>
    public DocxApplyResult Apply(
        Stream input,
        TextReader patchReader,
        Stream output,
        DocxEditOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(patchReader);
        ArgumentNullException.ThrowIfNull(output);
        if (!output.CanWrite)
        {
            throw new ArgumentException("Output stream must be writable.", nameof(output));
        }

        // The patch reader stays caller-owned. Input and output are disposed on
        // every exit below if and only if LeaveInputOpen/LeaveOutputOpen are
        // false (parsers, the loader, and disposal itself are idempotent, so
        // already-disposed streams are safe).
        try
        {
            // Revision metadata is normalized once here; the engine and every result below share the effective values.
            DocxEditOptions effectiveOptions = options.WithNormalizedRevisionMetadata();
            DocxPatch patch = ParsePatch(patchReader, options.MaxPatchChars, cancellationToken);
            if (!patch.Success)
            {
                return new DocxApplyResult
                {
                    Success = false,
                    Diagnostics = patch.Diagnostics,
                    Author = effectiveOptions.Author,
                    TrackChanges = effectiveOptions.TrackChanges,
                    TimestampUtc = effectiveOptions.TimestampUtc
                };
            }

            OoxmlPackage? package = TryLoad(input, ToPackageOptions(options.Quotas, options.LeaveInputOpen, allowMacroEnabledDocuments: options.AllowMacroEnabledDocuments), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
            if (package is null)
            {
                return new DocxApplyResult
                {
                    Success = false,
                    Diagnostics = diagnostics,
                    Author = effectiveOptions.Author,
                    TrackChanges = effectiveOptions.TrackChanges,
                    TimestampUtc = effectiveOptions.TimestampUtc
                };
            }

            if (!TryDocumentOperation(() => DocxPatchEngine.Apply(package, patch, effectiveOptions, cancellationToken), out PatchExecutionResult? execution, out IReadOnlyList<DocxDiagnostic> executionDiagnostics))
            {
                return new DocxApplyResult
                {
                    Success = false,
                    Diagnostics = diagnostics.Concat(executionDiagnostics).ToArray(),
                    Author = effectiveOptions.Author,
                    TrackChanges = effectiveOptions.TrackChanges,
                    TimestampUtc = effectiveOptions.TimestampUtc
                };
            }

            if (!execution.Success)
            {
                return new DocxApplyResult
                {
                    Success = false,
                    Diagnostics = diagnostics.Concat(execution.Diagnostics).ToArray(),
                    Operations = execution.Reports,
                    Author = effectiveOptions.Author,
                    TrackChanges = effectiveOptions.TrackChanges,
                    TimestampUtc = effectiveOptions.TimestampUtc
                };
            }

            package.Save(output, cancellationToken);
            return new DocxApplyResult
            {
                Success = diagnostics.All(d => d.Severity != DocxSeverity.Error),
                Diagnostics = diagnostics.Concat(execution.Diagnostics).ToArray(),
                Operations = execution.Reports,
                    Author = effectiveOptions.Author,
                    TrackChanges = effectiveOptions.TrackChanges,
                    TimestampUtc = effectiveOptions.TimestampUtc
            };
        }
        finally
        {
            if (!options.LeaveInputOpen)
            {
                input.Dispose();
            }

            if (!options.LeaveOutputOpen)
            {
                output.Dispose();
            }
        }
    }

    private static string[] SortedPartNames(OoxmlPackage package)
    {
        return package.Parts.Keys.Order(StringComparer.Ordinal).ToArray();
    }

    private static OoxmlPackage? TryLoad(
        Stream input,
        OoxmlPackageOptions options,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        if (!input.CanRead)
        {
            throw new ArgumentException("Input stream must be readable.", nameof(input));
        }

        try
        {
            OoxmlPackage package = OoxmlPackage.Load(input, options, cancellationToken);
            diagnostics = [];
            return package;
        }
        catch (Exception ex) when (IsExpectedDocumentException(ex))
        {
            diagnostics =
            [
                ToDocumentDiagnostic(ex)
            ];
            return null;
        }
    }

    private static bool TryDocumentOperation<T>(
        Func<T> operation,
        [NotNullWhen(true)] out T? result,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
        where T : class
    {
        try
        {
            result = operation();
            diagnostics = [];
            return true;
        }
        catch (Exception ex) when (IsExpectedDocumentException(ex))
        {
            result = null;
            diagnostics = [ToDocumentDiagnostic(ex)];
            return false;
        }
    }

    // Only document-origin exception types convert into failed results;
    // anything else still throws, so programmer errors are never masked here.
    private static bool IsExpectedDocumentException(Exception exception)
    {
        return exception is InvalidDataException or XmlException or IOException;
    }

    // Runs the unsupported-feature scan under the same document-exception
    // boundary as the primary scan: document-origin failures become failed
    // results with diagnostics, while programmer errors still throw.
    private static bool TryUnsupportedScan(
        OoxmlPackage package,
        bool includeHeadersFooters,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> unsupported,
        out IReadOnlyList<DocxDiagnostic> failure)
    {
        try
        {
            unsupported = DocxUnsupportedFeatureScanner.Scan(package, includeHeadersFooters, cancellationToken);
            failure = [];
            return true;
        }
        catch (Exception ex) when (IsExpectedDocumentException(ex))
        {
            unsupported = [];
            failure = [ToDocumentDiagnostic(ex)];
            return false;
        }
    }

    private static DocxDiagnostic ToDocumentDiagnostic(Exception exception)
    {
        if (exception is XmlException xmlException)
        {
            return new DocxDiagnostic(DocxSeverity.Error, "E0001", xmlException.Message) with
            {
                Line = xmlException.LineNumber > 0 ? xmlException.LineNumber : null,
                Column = xmlException.LinePosition > 0 ? xmlException.LinePosition : null
            };
        }

        return new DocxDiagnostic(DocxSeverity.Error, "E0001", exception.Message);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxPackageLimits quotas, bool leaveInputOpen, bool allowMacroEnabledDocuments)
    {
        return new OoxmlPackageOptions(
            leaveInputOpen,
            quotas.MaxZipEntries,
            quotas.MaxUncompressedBytes,
            quotas.MaxSinglePartBytes,
            allowMacroEnabledDocuments);
    }
}
