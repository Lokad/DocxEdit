using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

// D13: named result bindings for created objects. Creation operations accept
// an as field that names the new object; later operations address it through an at-sign name target. Bindings are recorded on snapshot alias marks, so
// resolution needs no alias table threading: paragraph and bookmark marks
// live in story parts, comment marks in comments parts, and all marks are
// stripped before publication. Bindings are local to one patch and resolve
// sequentially: duplicates, unknown names, forward references, and deleted
// targets fail explicitly, and an alias always names one created object.
internal static partial class DocxPatchEngine
{
    private static bool IsAliasReference(string target)
    {
        return target.StartsWith("@", StringComparison.Ordinal);
    }

    private static string AliasReferenceName(string target)
    {
        return target.Substring(1);
    }

    private static bool IsValidAliasName(string name)
    {
        if (name.Length == 0 || !char.IsLetter(name[0]))
        {
            return false;
        }

        foreach (char c in name)
        {
            if (!char.IsLetterOrDigit(c) && c != (char)95 && c != (char)45)
            {
                return false;
            }
        }

        return true;
    }

    private static (string PartName, XDocument Document, XElement Element)? FindAliasElement(
        OoxmlPackage package,
        string alias,
        CancellationToken cancellationToken)
    {
        foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(story.PartName);
            if (part is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, story.PartName, cancellationToken, out _);
            XElement? match = document.Descendants().FirstOrDefault(element =>
                (element.Name == OoxmlNs.W + "p" || element.Name == OoxmlNs.W + "bookmarkStart" || element.Name == OoxmlNs.W + "tr") &&
                string.Equals((string?)element.Attribute(SnapshotAliasName), alias, StringComparison.Ordinal));
            if (match is not null)
            {
                return (story.PartName, document, match);
            }
        }

        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            if (package.GetPart(partName) is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement? match = document.Descendants(OoxmlNs.W + "comment").FirstOrDefault(element =>
                string.Equals((string?)element.Attribute(SnapshotAliasName), alias, StringComparison.Ordinal));
            if (match is not null)
            {
                return (partName, document, match);
            }
        }

        return null;
    }

    private static DocxTargetId? TryResolveAliasParagraphId(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        string alias = AliasReferenceName(target);
        var found = FindMarkedStoryElement(package, SnapshotAliasName, alias, OoxmlNs.W + "p", cancellationToken);
        if (found is null)
        {
            return null;
        }

        int? ordinal = PhysicalParagraphOrdinal(found.Value.Document, found.Value.Element);
        if (ordinal is null)
        {
            return null;
        }

        (char storyLetter, int storyPart) = DocxTargetId.ParseStoryPrefix(found.Value.Story.Prefix);
        return new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Paragraph, ordinal.Value, 0, 0);
    }

    private static ParagraphTarget? ResolveAliasParagraphTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        (string PartName, XDocument Document, XElement Element)? found = FindAliasElement(package, AliasReferenceName(target), cancellationToken);
        if (found is null)
        {
            diagnostics = [Diagnostic(DocxSeverity.Error, "E1201", "Unknown result alias.", operation, target)];
            return null;
        }

        (string partName, XDocument document, XElement element) = found.Value;
        if (element.Name != OoxmlNs.W + "p")
        {
            diagnostics = [Diagnostic(DocxSeverity.Error, "E1201", "Result alias does not identify a paragraph.", operation, target)];
            return null;
        }

        return new ParagraphTarget(partName, document, element);
    }

    private static BlockTarget? ResolveAliasBlockTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        (string PartName, XDocument Document, XElement Element)? found = FindAliasElement(package, AliasReferenceName(target), cancellationToken);
        if (found is null)
        {
            diagnostics = [Diagnostic(DocxSeverity.Error, "E1201", "Unknown result alias.", operation, target)];
            return null;
        }

        (string partName, XDocument document, XElement element) = found.Value;
        if (element.Name != OoxmlNs.W + "p")
        {
            diagnostics = [Diagnostic(DocxSeverity.Error, "E1201", "Result alias does not identify a block.", operation, target)];
            return null;
        }

        return new BlockTarget(partName, document, element);
    }

    private static CommentTarget? ResolveAliasCommentTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out DocxDiagnostic? diagnostic)
    {
        diagnostic = null;
        (string PartName, XDocument Document, XElement Element)? found = FindAliasElement(package, AliasReferenceName(target), cancellationToken);
        if (found is null)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E1201", "Unknown result alias.", operation, target);
            return null;
        }

        (string partName, XDocument document, XElement element) = found.Value;
        if (element.Name != OoxmlNs.W + "comment")
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E1201", "Result alias does not identify a comment.", operation, target);
            return null;
        }

        return new CommentTarget(partName, document, element);
    }

    private static BookmarkTarget? ResolveAliasBookmarkTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        (string PartName, XDocument Document, XElement Element)? found = FindAliasElement(package, AliasReferenceName(target), cancellationToken);
        if (found is null)
        {
            diagnostics = [Diagnostic(DocxSeverity.Error, "E1201", "Unknown result alias.", operation, target)];
            return null;
        }

        (string partName, XDocument document, XElement start) = found.Value;
        if (start.Name != OoxmlNs.W + "bookmarkStart")
        {
            diagnostics = [Diagnostic(DocxSeverity.Error, "E1201", "Result alias does not identify a bookmark.", operation, target)];
            return null;
        }

        string? ooxmlId = (string?)start.Attribute(OoxmlNs.W + "id");
        XElement? end = ooxmlId is null
            ? null
            : document.Descendants(OoxmlNs.W + "bookmarkEnd").FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), ooxmlId, StringComparison.Ordinal));
        return new BookmarkTarget(partName, document, start, end);
    }

    private static HyperlinkTarget? ResolveAliasHyperlinkTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        string alias = AliasReferenceName(target);
        var found = FindMarkedStoryElement(package, SnapshotAliasName, alias, OoxmlNs.W + "hyperlink", cancellationToken);
        if (found is null)
        {
            return null;
        }

        return new HyperlinkTarget(found.Value.Story.PartName, found.Value.Document, found.Value.Element);
    }

    private static ImageBlipTarget? ResolveAliasImageBlipTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        string alias = AliasReferenceName(target);
        var found = FindMarkedStoryElement(package, SnapshotAliasName, alias, OoxmlNs.A + "blip", cancellationToken);
        if (found is null)
        {
            return null;
        }

        XDocument document = found.Value.Document;
        XElement match = found.Value.Element;
        string? relationshipId = (string?)match.Attribute(OoxmlNs.R + "embed");
        if (relationshipId is null)
        {
            return null;
        }
        OoxmlRelationship? relationship = package.GetRelationships(found.Value.Story.PartName, cancellationToken).FirstOrDefault(candidate => string.Equals(candidate.Id, relationshipId, StringComparison.Ordinal));
        if (relationship is null || relationship.IsExternal || relationship.ResolvedTarget is null)
        {
            return null;
        }
        OoxmlPart? part = package.GetPart(relationship.ResolvedTarget);
        if (part is null)
        {
            return null;
        }
        return new ImageBlipTarget(found.Value.Story.PartName, document, match, relationshipId, part);
    }

    private static BlockTarget? ResolveInsertAnchor(
        DocxPatchOperation operation,
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        if (target is null)
        {
            return null;
        }

        BlockTarget? anchor = ResolveBlockTarget(package, operation, target, cancellationToken, out _);
        if (anchor is null || (anchor.Block.Name != OoxmlNs.W + "p" && anchor.Block.Name != OoxmlNs.W + "tbl"))
        {
            return null;
        }

        return anchor;
    }

    private static List<XElement> FindAdjacentInsertParagraphs(XElement anchor, bool insertAfter, int count)
    {
        return insertAfter
            ? anchor.ElementsAfterSelf(OoxmlNs.W + "p").Take(count).ToList()
            : anchor.ElementsBeforeSelf(OoxmlNs.W + "p").TakeLast(count).ToList();
    }

    private static (string PartName, XDocument Document, XElement Element)? FindCommentElementById(
        OoxmlPackage package,
        string commentId,
        CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            if (package.GetPart(partName) is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement? match = document.Descendants(OoxmlNs.W + "comment").FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), commentId, StringComparison.Ordinal));
            if (match is not null)
            {
                return (partName, document, match);
            }
        }

        return null;
    }


    private static RowTarget? ResolveAliasRowTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        (string PartName, XDocument Document, XElement Element)? found = FindAliasElement(package, AliasReferenceName(target), cancellationToken);
        if (found is null)
        {
            return null;
        }

        (string partName, XDocument document, XElement row) = found.Value;
        if (row.Name != OoxmlNs.W + "tr")
        {
            return null;
        }

        XElement? table = row.Parent?.Name == OoxmlNs.W + "tbl" ? row.Parent : row.Ancestors(OoxmlNs.W + "tbl").FirstOrDefault();
        return table is null ? null : new RowTarget(partName, document, table, row);
    }

    private static (string PartName, XDocument Document, XElement Row)? FindCreatedRow(
        DocxPatchOperation operation,
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        if (target is null)
        {
            return null;
        }

        if (operation.OperationName == "append-row")
        {
            TableTarget? tableTarget = ResolveTableTarget(package, target, cancellationToken);
            if (tableTarget is null)
            {
                return null;
            }

            XElement? created = tableTarget.Table.Elements(OoxmlNs.W + "tr").LastOrDefault();
            return created is null ? null : (tableTarget.PartName, tableTarget.Document, created);
        }

        RowTarget? anchor = ResolveRowTarget(package, target, cancellationToken);
        if (anchor is null)
        {
            return null;
        }

        bool insertAfter = string.Equals(operation.OperationName, "insert-row-after", StringComparison.Ordinal);
        XElement? sibling = insertAfter
            ? anchor.Row.ElementsAfterSelf(OoxmlNs.W + "tr").FirstOrDefault()
            : anchor.Row.ElementsBeforeSelf(OoxmlNs.W + "tr").LastOrDefault();
        if (sibling is null)
        {
            return null;
        }

        XElement? table = sibling.Parent?.Name == OoxmlNs.W + "tbl" ? sibling.Parent : sibling.Ancestors(OoxmlNs.W + "tbl").FirstOrDefault();
        return table is null ? null : (anchor.PartName, anchor.Document, sibling);
    }

    private static bool BindCreatedAlias(
        DocxPatchOperation operation,
        OoxmlPackage package,
        string alias,
        IReadOnlySet<string>? commentsBefore,
        IReadOnlyDictionary<string, HashSet<string>>? bookmarkIdsBefore,
        CancellationToken cancellationToken)
    {
        Dictionary<string, XDocument> touched = new(StringComparer.OrdinalIgnoreCase);

        var aliasPrefixes = DocxPartRoles.GetStoryPrefixes(package, cancellationToken);
        foreach (var aliasEntry in aliasPrefixes)
        {
            XDocument aliasDocument = LoadDocumentPart(package, aliasEntry.Key, cancellationToken, out XElement _);
            foreach (XElement aliasElement in aliasDocument.Descendants())
            {
                if (string.Equals((string?)aliasElement.Attribute(SnapshotAliasName), alias, StringComparison.Ordinal))
                {
                    return true;
                }
            }
        }
        if (operation.OperationName is "insert-before" or "insert-after" or "insert-image-after" or "insert-hyperlink-after")
        {
            BlockTarget? anchor = ResolveInsertAnchor(operation, package, cancellationToken);
            if (anchor is null)
            {
                return false;
            }

            bool insertAfter = !string.Equals(operation.OperationName, "insert-before", StringComparison.Ordinal);
            int count = Math.Max(1, operation.FieldValues.Count(static field => field.Name == "text"));
            foreach (XElement created in FindAdjacentInsertParagraphs(anchor.Block, insertAfter, count))
            {
                created.SetAttributeValue(SnapshotAliasName, alias);
                if (string.Equals(operation.OperationName, "insert-hyperlink-after", StringComparison.Ordinal))
                {
                    created.Descendants(OoxmlNs.W + "hyperlink").FirstOrDefault()?.SetAttributeValue(SnapshotAliasName, alias);
                }
                if (string.Equals(operation.OperationName, "insert-image-after", StringComparison.Ordinal))
                {
                    created.Descendants(OoxmlNs.A + "blip").FirstOrDefault()?.SetAttributeValue(SnapshotAliasName, alias);
                }
            }

            touched[anchor.PartName] = anchor.Document;
            SaveTouchedParts(package, touched);
            return true;
        }

        if ((operation.OperationName == "add-comment" || operation.OperationName == "add-comment-reply") && commentsBefore is not null)
        {
            HashSet<string> after = ReadCommentIds(package, cancellationToken);
            after.ExceptWith(commentsBefore);
            bool boundAlias = false;
            foreach (string id in after)
            {
                (string PartName, XDocument Document, XElement Element)? found = FindCommentElementById(package, id, cancellationToken);
                if (found is null)
                {
                    continue;
                }

                found.Value.Element.SetAttributeValue(SnapshotAliasName, alias);
                boundAlias = true;
                touched[found.Value.PartName] = found.Value.Document;
            }

            SaveTouchedParts(package, touched);
            return boundAlias;
        }

        if (operation.OperationName == "add-bookmark" && bookmarkIdsBefore is not null)
        {
            Dictionary<string, HashSet<string>> after = CollectBookmarkOoxmlIds(package, cancellationToken);
            bool bookmarkBound = false;
            foreach ((string partName, HashSet<string> ids) in after)
            {
                bookmarkIdsBefore.TryGetValue(partName, out HashSet<string>? before);
                List<string> added = ids.Where(id => before is null || !before.Contains(id)).ToList();
                if (added.Count != 1)
                {
                    continue;
                }

                if (package.GetPart(partName) is null)
                {
                    continue;
                }

                XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
                XElement? start = document.Descendants(OoxmlNs.W + "bookmarkStart").FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), added[0], StringComparison.Ordinal));
                if (start is null)
                {
                    continue;
                }

                start.SetAttributeValue(SnapshotAliasName, alias);
                bookmarkBound = true;
                touched[partName] = document;
            }

            SaveTouchedParts(package, touched);
            return bookmarkBound;
        }

        if (operation.OperationName is "append-row" or "insert-row-before" or "insert-row-after")
        {
            (string PartName, XDocument Document, XElement Row)? found = FindCreatedRow(operation, package, cancellationToken);
            if (found is null)
            {
                return false;
            }

            found.Value.Row.SetAttributeValue(SnapshotAliasName, alias);
            touched[found.Value.PartName] = found.Value.Document;
            SaveTouchedParts(package, touched);
            return true;
        }
        return false;
    }

    private static void SaveTouchedParts(OoxmlPackage package, Dictionary<string, XDocument> touched)
    {
        foreach ((string partName, XDocument document) in touched)
        {
            SaveDocumentPart(package, partName, document);
        }
    }

    private static DocxDiagnostic? ValidateAliasDefinition(
        DocxPatchOperation operation,
        IReadOnlySet<string> definedAliases)
    {
        if (!operation.Fields.TryGetValue("as", out string? alias) || alias is null)
        {
            return null;
        }

        string? target = operation.Fields.GetValueOrDefault("target");
        if (!IsValidAliasName(alias))
        {
            return Diagnostic(DocxSeverity.Error, "E4205", "Invalid alias name: use a leading letter followed by letters, digits, underscore, or hyphen.", operation, target, fieldName: "as");
        }

        if (definedAliases.Contains(alias))
        {
            return Diagnostic(DocxSeverity.Error, "E4205", "Duplicate alias: this patch already binds that name.", operation, target, fieldName: "as");
        }

        if ((operation.OperationName is "insert-before" or "insert-after" or "insert-image-after") &&
            operation.FieldValues.Count(static field => field.Name == "text") > 1)
        {
            return Diagnostic(DocxSeverity.Error, "E4205", "An alias requires a single created object: use one text field with as.", operation, target, fieldName: "as");
        }

        return null;
    }

    internal static IReadOnlyList<DocxDiagnostic> ValidatePatchAliases(DocxPatch patch)
    {
        List<DocxDiagnostic> diagnostics = [];
        HashSet<string> defined = new(StringComparer.Ordinal);
        foreach (DocxPatchOperation operation in patch.Operations)
        {
            if (operation.Fields.TryGetValue("as", out string? alias) && alias is not null)
            {
                string? aliasTarget = operation.Fields.GetValueOrDefault("target");
                if (!IsValidAliasName(alias))
                {
                    diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Invalid alias name: use a leading letter followed by letters, digits, underscore, or hyphen.", operation, aliasTarget, fieldName: "as"));
                }
                else if (defined.Contains(alias))
                {
                    diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Duplicate alias: this patch already binds that name.", operation, aliasTarget, fieldName: "as"));
                }
                else
                {
                    defined.Add(alias);
                    if ((operation.OperationName is "insert-before" or "insert-after" or "insert-image-after") &&
                        operation.FieldValues.Count(static field => field.Name == "text") > 1)
                    {
                        diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "An alias requires a single created object: use one text field with as.", operation, aliasTarget, fieldName: "as"));
                    }
                }
            }

            string? target = operation.Fields.GetValueOrDefault("target");
            if (target is not null && IsAliasReference(target) && !defined.Contains(AliasReferenceName(target)))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E1201", "Unknown result alias.", operation, target));
            }
        }

        return diagnostics;
    }
}
