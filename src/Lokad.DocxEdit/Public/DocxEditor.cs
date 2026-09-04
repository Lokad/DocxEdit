using Lokad.DocxEdit.Ooxml;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Rendering;
using System.Xml;

namespace Lokad.DocxEdit;

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

    public DocxReadResult Read(
        Stream input,
        DocxReadOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxReadResult
            {
                Success = false,
                Diagnostics = diagnostics
            };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, options.IncludeHeadersFooters || options.IncludeAllStories, options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxReadResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                PartNames = SortedPartNames(package),
                MainDocumentPartName = package.MainDocumentPartName
            };
        }

        return new DocxReadResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, options.IncludeHeadersFooters || options.IncludeAllStories, cancellationToken))
                .ToArray(),
            PartNames = SortedPartNames(package),
            MainDocumentPartName = package.MainDocumentPartName,
            Text = TextRenderers.RenderRead(model!, options.MaxText),
            Paragraphs = model!.Paragraphs,
            Tables = model.Tables,
            Images = model.Images,
            Sections = model.Sections,
            Bookmarks = model.Bookmarks,
            ContentControls = model.ContentControls,
            Fields = model.Fields,
            Hyperlinks = model.Hyperlinks
        };
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

    public DocxOutlineResult Outline(
        Stream input,
        DocxOutlineOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
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

        return new DocxOutlineResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, options.IncludeHeadersFooters, cancellationToken))
                .ToArray(),
            PartNames = SortedPartNames(package),
            MainDocumentPartName = package.MainDocumentPartName,
            Lines = TextRenderers.RenderOutline(model!)
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

    public DocxFindResult Find(
        Stream input,
        string query,
        DocxFindOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(query);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
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

        return new DocxFindResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, options.IncludeHeadersFooters, cancellationToken))
                .ToArray(),
            Query = query,
            Matches = TextRenderers.Find(model!, query, options.MaxText)
        };
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

    public DocxDumpResult Dump(
        Stream input,
        string targetId,
        DocxDumpOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetId);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxDumpResult { Success = false, Diagnostics = diagnostics, TargetId = targetId };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, includeHeadersFooters: false, textView: options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
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

        return new DocxDumpResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, includeHeadersFooters: false, cancellationToken))
                .ToArray(),
            TargetId = targetId,
            Text = TextRenderers.Dump(model!, changes!, targetId, options.IncludeRuns, options.MaxText),
            Runs = options.IncludeRuns ? TextRenderers.DumpRuns(model!, targetId, options.MaxText) : []
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

    public DocxContextResult Context(
        Stream input,
        string targetId,
        DocxContextOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetId);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxContextResult { Success = false, Diagnostics = diagnostics, TargetId = targetId };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, options.IncludeHeadersFooters, options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
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

        IReadOnlyList<DocxContextItem> items = TextRenderers.Context(model!, changes!, targetId, options.Radius, options.MaxText);
        IReadOnlyList<DocxDiagnostic> contextDiagnostics = items.Count == 0
            ? [new DocxDiagnostic(DocxSeverity.Error, "E2001", $"Target '{targetId}' was not found.", TargetId: targetId)]
            : [];
        return new DocxContextResult
        {
            Success = items.Count > 0,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, options.IncludeHeadersFooters, cancellationToken))
                .Concat(contextDiagnostics)
                .ToArray(),
            TargetId = targetId,
            Items = items,
            Text = TextRenderers.RenderContext(items)
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

    public DocxStylesResult Styles(
        Stream input,
        DocxStylesOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
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
            Styles = styles!
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

    public DocxMediaResult Media(
        Stream input,
        DocxMediaOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
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

        return new DocxMediaResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, includeHeadersFooters: false, cancellationToken))
                .ToArray(),
            Images = model!.Images
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

    public DocxValidateResult Validate(
        Stream input,
        DocxValidateOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
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
            diagnostics.Concat(validationDiagnostics!),
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

    public DocxChangesResult Changes(
        Stream input,
        DocxChangesOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: false), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
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

        IReadOnlyList<DocxChangeInfo> annotatedChanges = AnnotateOperationReports(changes!, options.OperationReports);

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

    /// <summary>Parses a patch. Uses default options and no cancellation.</summary>
    public DocxPatch ParsePatch(
        TextReader patchReader)
    {
        return ParsePatch(patchReader, new DocxPatchParseOptions(), CancellationToken.None);
    }

    /// <summary>Parses a patch. Uses no cancellation.</summary>
    public DocxPatch ParsePatch(
        TextReader patchReader, DocxPatchParseOptions options)
    {
        return ParsePatch(patchReader, options, CancellationToken.None);
    }

    public DocxPatch ParsePatch(
        TextReader patchReader,
        DocxPatchParseOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(patchReader);
        _ = options;
        cancellationToken.ThrowIfCancellationRequested();
        string text = patchReader.ReadToEnd();
        cancellationToken.ThrowIfCancellationRequested();
        return DocxPatchParser.Parse(text);
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

    public DocxCheckResult Check(
        Stream input,
        TextReader patchReader,
        DocxEditOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(patchReader);

        DocxPatch patch = ParsePatch(patchReader, new DocxPatchParseOptions(), cancellationToken);
        if (!patch.Success)
        {
            return new DocxCheckResult { Success = false, Diagnostics = patch.Diagnostics };
        }

        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: options.AllowMacroEnabledDocuments), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxCheckResult
            {
                Success = false,
                Diagnostics = diagnostics
            };
        }

        if (!TryDocumentOperation(() => DocxPatchEngine.Check(package, patch, options, cancellationToken), out PatchExecutionResult? execution, out IReadOnlyList<DocxDiagnostic> executionDiagnostics))
        {
            return new DocxCheckResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(executionDiagnostics).ToArray()
            };
        }

        return new DocxCheckResult
        {
            Success = execution!.Success,
            Diagnostics = diagnostics.Concat(execution.Diagnostics).ToArray(),
            Operations = execution.Reports
        };
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

        DocxPatch patch = ParsePatch(patchReader, new DocxPatchParseOptions(), cancellationToken);
        if (!patch.Success)
        {
            return new DocxApplyResult { Success = false, Diagnostics = patch.Diagnostics };
        }

        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options, allowMacroEnabledDocuments: options.AllowMacroEnabledDocuments), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxApplyResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxPatchEngine.Apply(package, patch, options, cancellationToken), out PatchExecutionResult? execution, out IReadOnlyList<DocxDiagnostic> executionDiagnostics))
        {
            return new DocxApplyResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(executionDiagnostics).ToArray()
            };
        }

        if (!execution!.Success)
        {
            return new DocxApplyResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(execution.Diagnostics).ToArray(),
                Operations = execution.Reports
            };
        }

        package.Save(output, cancellationToken);
        if (!options.LeaveOutputOpen)
        {
            output.Dispose();
        }

        return new DocxApplyResult
        {
            Success = diagnostics.All(d => d.Severity != DocxSeverity.Error),
            Diagnostics = diagnostics.Concat(execution.Diagnostics).ToArray(),
            Operations = execution.Reports
        };
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
        out T? result,
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

    private static bool IsExpectedDocumentException(Exception exception)
    {
        return exception is InvalidDataException or XmlException or IOException;
    }

    private static DocxDiagnostic ToDocumentDiagnostic(Exception exception)
    {
        if (exception is XmlException xmlException)
        {
            return new DocxDiagnostic(
                DocxSeverity.Error,
                "E0001",
                xmlException.Message,
                Line: xmlException.LineNumber > 0 ? xmlException.LineNumber : null,
                Column: xmlException.LinePosition > 0 ? xmlException.LinePosition : null);
        }

        return new DocxDiagnostic(DocxSeverity.Error, "E0001", exception.Message);
    }

    private static OoxmlPackageOptions ToPackageOptions(IHasPackageLimits options, bool allowMacroEnabledDocuments)
    {
        return new OoxmlPackageOptions(
            options.LeaveInputOpen,
            options.MaxZipEntries,
            options.MaxUncompressedBytes,
            options.MaxSinglePartBytes,
            allowMacroEnabledDocuments);
    }
}
