using System.Globalization;
using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxDocumentScanner
{
    private static IReadOnlyList<DocxBookmarkInfo> ReadBookmarks(
        XDocument document,
        string partName,
        string story,
        string idPrefix,
        IReadOnlyDictionary<XElement, string> targets)
    {
        var bookmarks = new List<DocxBookmarkInfo>();
        var endsByOoxmlId = document
            .Descendants(OoxmlNs.W + "bookmarkEnd")
            .WithNonBlankKey(end => (string?)end.Attribute(OoxmlNs.W + "id"))
            .GroupBy(pair => pair.Key, pair => pair.Item, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

        int bookmarkIndex = 1;
        foreach (XElement start in document.Descendants(OoxmlNs.W + "bookmarkStart"))
        {
            string? name = (string?)start.Attribute(OoxmlNs.W + "name");
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            string? ooxmlId = (string?)start.Attribute(OoxmlNs.W + "id");
            endsByOoxmlId.TryGetValue(ooxmlId ?? string.Empty, out XElement? end);
            bookmarks.Add(new DocxBookmarkInfo
            {
                Id = $"{idPrefix}.B{bookmarkIndex++:0000}",
                Name = name,
                OoxmlId = ooxmlId,
                Story = story,
                PartName = partName,
                StartTargetId = FindTargetId(start, targets),
                EndTargetId = end is null ? null : FindTargetId(end, targets),
                IsComplete = end is not null
            });
        }

        return AnnotateDuplicateBookmarkNames(bookmarks);
    }

    private static IReadOnlyList<DocxBookmarkInfo> AnnotateDuplicateBookmarkNames(IReadOnlyList<DocxBookmarkInfo> bookmarks)
    {
        IReadOnlyDictionary<string, string[]> duplicateIdsByName = BuildDuplicateIds(
            bookmarks,
            bookmark => bookmark.Name,
            bookmark => bookmark.Id);
        return bookmarks
            .Select(bookmark => duplicateIdsByName.TryGetValue(bookmark.Name, out string[]? ids)
                ? bookmark with { IsNameDuplicate = true, DuplicateNameBookmarkIds = ids }
                : bookmark)
            .ToArray();
    }

    private static IReadOnlyList<DocxContentControlInfo> ReadContentControls(
        XDocument document,
        string partName,
        string story,
        string idPrefix,
        DocxTextView textView,
        IReadOnlyDictionary<XElement, string> targets)
    {
        var contentControls = new List<ContentControlScanEntry>();
        int controlIndex = 1;
        foreach (XElement control in document.Descendants(OoxmlNs.W + "sdt"))
        {
            XElement? properties = control.Element(OoxmlNs.W + "sdtPr");
            XElement content = control.Element(OoxmlNs.W + "sdtContent") ?? control;
            string kind = ReadContentControlKind(properties);
            string? lockValue = ReadSdtLock(properties);
            contentControls.Add(new ContentControlScanEntry(control, new DocxContentControlInfo
            {
                Id = $"{idPrefix}.CC{controlIndex++:0000}",
                Story = story,
                PartName = partName,
                TargetId = FindTargetId(control, targets),
                Kind = kind,
                OoxmlId = ReadSdtProperty(properties, "id"),
                Tag = ReadSdtProperty(properties, "tag"),
                Alias = ReadSdtProperty(properties, "alias"),
                PlaceholderDocPart = ReadPlaceholderDocPart(properties),
                IsShowingPlaceholderText = properties?.Element(OoxmlNs.W + "showingPlcHdr") is not null,
                DataBindingXPath = ReadSdtAttribute(properties, "dataBinding", "xpath"),
                DataBindingStoreItemId = ReadSdtAttribute(properties, "dataBinding", "storeItemID"),
                DataBindingPrefixMappings = ReadSdtAttribute(properties, "dataBinding", "prefixMappings"),
                RepeatingSectionTitle = ReadSdtAttribute(properties, "repeatingSection", "sectionTitle"),
                RepeatingSectionItemCount = kind == "repeating-section"
                    ? content.Elements(OoxmlNs.W + "sdt").Count(child => child.Element(OoxmlNs.W + "sdtPr")?.Element(OoxmlNs.W + "repeatingSectionItem") is not null)
                    : null,
                SafeEditStatus = ReadContentControlSafeEditStatus(kind, lockValue),
                SafeEditReason = ReadContentControlSafeEditReason(kind, lockValue),
                Lock = lockValue,
                Checked = ReadContentControlChecked(properties),
                CheckedSymbol = ReadContentControlStateSymbol(properties, "checkedState"),
                UncheckedSymbol = ReadContentControlStateSymbol(properties, "uncheckedState"),
                ListItems = ReadContentControlListItems(properties),
                DateFormat = ReadNestedSdtProperty(properties, "date", "dateFormat"),
                DateLanguage = ReadNestedSdtProperty(properties, "date", "lid"),
                DateCalendar = ReadNestedSdtProperty(properties, "date", "calendar"),
                DateValue = ReadNestedSdtProperty(properties, "date", "fullDate"),
                TextLength = ReadText(content, textView).Length
            }));
        }

        return AnnotateContentControlHierarchy(AnnotateDuplicateContentControlSelectors(contentControls.Select(entry => entry.Info).ToArray()), contentControls);
    }

    private static IReadOnlyList<DocxContentControlInfo> AnnotateDuplicateContentControlSelectors(IReadOnlyList<DocxContentControlInfo> contentControls)
    {
        IReadOnlyDictionary<string, string[]> duplicateIdsByTag = BuildDuplicateIds(
            contentControls,
            control => control.Tag,
            control => control.Id);
        IReadOnlyDictionary<string, string[]> duplicateIdsByAlias = BuildDuplicateIds(
            contentControls,
            control => control.Alias,
            control => control.Id);
        return contentControls
            .Select(control =>
            {
                DocxContentControlInfo annotated = control;
                if (control.Tag is not null && duplicateIdsByTag.TryGetValue(control.Tag, out string[]? tagIds))
                {
                    annotated = annotated with { IsTagDuplicate = true, DuplicateTagControlIds = tagIds };
                }

                if (control.Alias is not null && duplicateIdsByAlias.TryGetValue(control.Alias, out string[]? aliasIds))
                {
                    annotated = annotated with { IsAliasDuplicate = true, DuplicateAliasControlIds = aliasIds };
                }

                return annotated;
            })
            .ToArray();
    }

    private static IReadOnlyList<DocxContentControlInfo> AnnotateContentControlHierarchy(
        IReadOnlyList<DocxContentControlInfo> contentControls,
        IReadOnlyList<ContentControlScanEntry> entries)
    {
        IReadOnlyDictionary<XElement, string> idsByElement = entries.ToDictionary(entry => entry.Element, entry => entry.Info.Id);
        IReadOnlyDictionary<string, DocxContentControlInfo> controlsById = contentControls.ToDictionary(control => control.Id, StringComparer.Ordinal);
        return entries.Select(entry =>
        {
            DocxContentControlInfo info = controlsById[entry.Info.Id];
            XElement? parent = entry.Element.Ancestors(OoxmlNs.W + "sdt").FirstOrDefault(idsByElement.ContainsKey);
            string? parentId = parent is null ? null : idsByElement[parent];
            string[] childIds = entry.Element
                .Descendants(OoxmlNs.W + "sdt")
                .Where(child => child.Ancestors(OoxmlNs.W + "sdt").FirstOrDefault() == entry.Element)
                .Where(idsByElement.ContainsKey)
                .Select(child => idsByElement[child])
                .ToArray();
            return info with
            {
                ParentContentControlId = parentId,
                ChildContentControlIds = childIds
            };
        }).ToArray();
    }

    private static IReadOnlyDictionary<string, string[]> BuildDuplicateIds<T>(
        IEnumerable<T> items,
        Func<T, string?> keySelector,
        Func<T, string> idSelector)
    {
        return items
            .WithNonBlankKey(keySelector)
            .GroupBy(pair => pair.Key, pair => pair.Item, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .ToDictionary(
                group => group.Key,
                group => group.Select(item => idSelector(item)).ToArray(),
                StringComparer.Ordinal);
    }

    private static IReadOnlyList<DocxFieldInfo> ReadFields(
        XDocument document,
        string partName,
        string story,
        string idPrefix,
        DocxTextView textView,
        IReadOnlyDictionary<XElement, string> targets)
    {
        var fields = new List<DocxFieldInfo>();
        var stack = new Stack<ComplexFieldBuilder>();
        int fieldIndex = 1;
        foreach (XElement element in document.Descendants())
        {
            if (element.Name == OoxmlNs.W + "fldSimple")
            {
                string code = NormalizeFieldCode((string?)element.Attribute(OoxmlNs.W + "instr") ?? string.Empty);
                FieldCodeMetadata metadata = AnalyzeFieldCode(code);
                bool? isLocked = ReadOnOffAttribute(element, "fldLock");
                string cachedResultText = ReadText(element, textView);
                fields.Add(new DocxFieldInfo
                {
                    Id = $"{idPrefix}.F{fieldIndex++:0000}",
                    Story = story,
                    PartName = partName,
                    TargetId = FindTargetId(element, targets),
                    Kind = "simple",
                    FieldType = metadata.FieldType,
                    Code = code,
                    Arguments = metadata.Arguments,
                    Switches = metadata.Switches,
                    CachedResultText = cachedResultText,
                    ResultTextLength = cachedResultText.Length,
                    NestingDepth = stack.Count,
                    BookmarkDependencies = metadata.BookmarkDependencies,
                    HyperlinkDependencies = metadata.HyperlinkDependencies,
                    RefreshPolicy = metadata.RefreshPolicy,
                    RefreshReason = metadata.RefreshReason,
                    CanRefreshDeterministically = metadata.CanRefreshDeterministically,
                    SafeEditStatus = DetermineFieldSafeEditStatus("simple", complete: true, isLocked),
                    IsDirty = ReadOnOffAttribute(element, "dirty"),
                    IsLocked = isLocked,
                    IsComplete = true
                });
                continue;
            }

            if (element.Name == OoxmlNs.W + "fldChar")
            {
                string? fieldCharType = (string?)element.Attribute(OoxmlNs.W + "fldCharType");
                if (string.Equals(fieldCharType, "begin", StringComparison.Ordinal))
                {
                    stack.Push(new ComplexFieldBuilder(element, FindTargetId(element, targets), stack.Count)
                    {
                        IsDirty = ReadOnOffAttribute(element, "dirty"),
                        IsLocked = ReadOnOffAttribute(element, "fldLock")
                    });
                }
                else if (string.Equals(fieldCharType, "separate", StringComparison.Ordinal))
                {
                    if (stack.Count > 0)
                    {
                        stack.Peek().HasSeparate = true;
                    }
                }
                else if (string.Equals(fieldCharType, "end", StringComparison.Ordinal) && stack.Count > 0)
                {
                    ComplexFieldBuilder builder = stack.Pop();
                    fields.Add(builder.ToInfo($"{idPrefix}.F{fieldIndex++:0000}", story, partName, complete: true));
                }

                continue;
            }

            if (stack.Count == 0)
            {
                continue;
            }

            ComplexFieldBuilder current = stack.Peek();
            if (element.Name == OoxmlNs.W + "instrText")
            {
                current.Code.Append(element.Value);
            }
            else
            {
                AddComplexFieldResultText(stack, element, textView);
            }
        }

        foreach (ComplexFieldBuilder builder in stack)
        {
            fields.Add(builder.ToInfo($"{idPrefix}.F{fieldIndex++:0000}", story, partName, complete: false));
        }

        return fields;
    }

    private static void AddComplexFieldResultText(
        Stack<ComplexFieldBuilder> stack,
        XElement element,
        DocxTextView textView)
    {
        string text = ReadFieldResultTextElement(element, textView);
        if (text.Length == 0)
        {
            return;
        }

        foreach (ComplexFieldBuilder builder in stack)
        {
            if (builder.HasSeparate)
            {
                builder.ResultText.Append(text);
            }
        }
    }

    private static IReadOnlyList<DocxHyperlinkInfo> ReadHyperlinks(
        XDocument document,
        string partName,
        string story,
        string idPrefix,
        DocxTextView textView,
        IReadOnlyDictionary<XElement, string> targets,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships)
    {
        var hyperlinks = new List<DocxHyperlinkInfo>();
        IReadOnlyDictionary<string, int> bookmarkNameCounts = document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Select(bookmark => (string?)bookmark.Attribute(OoxmlNs.W + "name"))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .OfType<string>()
            .GroupBy(name => name, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        int hyperlinkIndex = 1;
        foreach (XElement hyperlink in document.Descendants(OoxmlNs.W + "hyperlink"))
        {
            string? relationshipId = (string?)hyperlink.Attribute(OoxmlNs.R + "id");
            relationships.TryGetValue(relationshipId ?? string.Empty, out OoxmlRelationship? relationship);
            string? anchor = (string?)hyperlink.Attribute(OoxmlNs.W + "anchor");
            string? externalUri = relationship?.IsExternal == true ? relationship.Target : null;
            HyperlinkUriValidation uriValidation = ValidateHyperlinkUri(externalUri);
            int anchorCount = !string.IsNullOrWhiteSpace(anchor) && bookmarkNameCounts.TryGetValue(anchor, out int resolvedAnchorCount)
                ? resolvedAnchorCount
                : 0;
            hyperlinks.Add(new DocxHyperlinkInfo
            {
                Id = $"{idPrefix}.L{hyperlinkIndex++:0000}",
                Story = story,
                PartName = partName,
                TargetId = FindTargetId(hyperlink, targets),
                RelationshipId = relationshipId,
                RelationshipPartName = string.IsNullOrWhiteSpace(relationshipId) ? null : OoxmlPath.GetRelationshipPartName(partName),
                RelationshipTargetMode = relationship?.TargetMode,
                Uri = externalUri,
                UriScheme = uriValidation.Scheme,
                IsUriValid = uriValidation.IsValid,
                UriValidationReason = uriValidation.Reason,
                Anchor = anchor,
                IsAnchorMissing = string.IsNullOrWhiteSpace(anchor) ? null : anchorCount == 0,
                IsAnchorDuplicate = string.IsNullOrWhiteSpace(anchor) ? null : anchorCount > 1,
                Tooltip = (string?)hyperlink.Attribute(OoxmlNs.W + "tooltip"),
                TargetFrame = (string?)hyperlink.Attribute(OoxmlNs.W + "tgtFrame"),
                History = ReadOnOffAttribute(hyperlink, "history"),
                TargetPartName = relationship?.IsExternal == false ? relationship.ResolvedTarget : null,
                IsExternal = relationship?.IsExternal == true,
                IsBroken = !string.IsNullOrWhiteSpace(relationshipId) && relationship is null,
                DisplayTextLength = ReadText(hyperlink, textView).Length
            });
        }

        return hyperlinks;
    }

    private static HyperlinkUriValidation ValidateHyperlinkUri(string? uri)
    {
        if (string.IsNullOrWhiteSpace(uri))
        {
            return HyperlinkUriValidation.None;
        }

        if (!Uri.TryCreate(uri, UriKind.RelativeOrAbsolute, out Uri? parsed))
        {
            return new HyperlinkUriValidation(null, false, "malformed-uri");
        }

        if (!parsed.IsAbsoluteUri)
        {
            return new HyperlinkUriValidation(null, false, "relative-uri");
        }

        string scheme = parsed.Scheme;
        bool valid = scheme is "http" or "https" or "mailto";
        return new HyperlinkUriValidation(
            scheme,
            valid,
            valid ? null : "unsupported-uri-scheme");
    }

    private static string? FindTargetId(XElement element, IReadOnlyDictionary<XElement, string> targets)
    {
        foreach (XElement candidate in element.AncestorsAndSelf())
        {
            if (targets.TryGetValue(candidate, out string? id))
            {
                return id;
            }
        }

        foreach (XElement descendant in element.Descendants())
        {
            if (targets.TryGetValue(descendant, out string? id))
            {
                return id;
            }
        }

        return null;
    }

    private static string ReadFieldResultTextElement(XElement element, DocxTextView textView)
    {
        if (!ShouldIncludeTextElement(element, textView))
        {
            return string.Empty;
        }

        if (element.Name == OoxmlNs.W + "t" || element.Name == OoxmlNs.W + "delText")
        {
            return ApplyMarkupTextView(element, element.Value, textView);
        }

        if (element.Name == OoxmlNs.W + "tab" || element.Name == OoxmlNs.W + "br")
        {
            return ApplyMarkupTextView(element, element.Name == OoxmlNs.W + "tab" ? "\t" : "\n", textView);
        }

        return string.Empty;
    }

    private static string NormalizeFieldCode(string code)
    {
        return string.Join(
            " ",
            code.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
    }

    private static FieldCodeMetadata AnalyzeFieldCode(string code)
    {
        string[] tokens = TokenizeFieldCode(code);
        if (tokens.Length == 0)
        {
            return new FieldCodeMetadata(null, [], [], [], [], "unsupported", "field code is empty", false);
        }

        string fieldType = NormalizeFieldType(tokens[0]);
        string[] arguments = ReadFieldArguments(tokens);
        string[] switches = ReadFieldSwitches(tokens);
        var bookmarkDependencies = new List<string>();
        var hyperlinkDependencies = new List<string>();
        if (fieldType is "REF" or "PAGEREF" or "NOTEREF")
        {
            AddFirstFieldOperand(tokens, bookmarkDependencies);
        }
        else if (fieldType == "HYPERLINK")
        {
            ReadHyperlinkFieldDependencies(tokens, hyperlinkDependencies, bookmarkDependencies);
        }

        FieldRefreshProfile refresh = ClassifyFieldRefresh(fieldType);
        return new FieldCodeMetadata(
            fieldType,
            arguments,
            switches,
            DistinctNonEmpty(bookmarkDependencies),
            DistinctNonEmpty(hyperlinkDependencies),
            refresh.Policy,
            refresh.Reason,
            refresh.CanRefreshDeterministically);
    }

    private static string[] TokenizeFieldCode(string code)
    {
        var tokens = new List<string>();
        var current = new System.Text.StringBuilder();
        bool inQuote = false;
        foreach (char character in code)
        {
            if (character == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (char.IsWhiteSpace(character) && !inQuote)
            {
                AddCurrentFieldToken(tokens, current);
                continue;
            }

            current.Append(character);
        }

        AddCurrentFieldToken(tokens, current);
        return tokens.ToArray();
    }

    private static void AddCurrentFieldToken(List<string> tokens, System.Text.StringBuilder current)
    {
        if (current.Length == 0)
        {
            return;
        }

        tokens.Add(current.ToString());
        current.Clear();
    }

    private static string NormalizeFieldType(string token)
    {
        return token.StartsWith('=') ? "FORMULA" : token.ToUpperInvariant();
    }

    private static string[] ReadFieldArguments(string[] tokens)
    {
        return tokens
            .Skip(1)
            .Where(token => !token.StartsWith('\\'))
            .ToArray();
    }

    private static string[] ReadFieldSwitches(string[] tokens)
    {
        return tokens
            .Skip(1)
            .Where(token => token.StartsWith('\\'))
            .Select(token => token.ToUpperInvariant())
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }

    private static FieldRefreshProfile ClassifyFieldRefresh(string fieldType)
    {
        return fieldType switch
        {
            "REF" or "PAGEREF" or "NOTEREF" => new("same-part-bookmark", "refreshes from an unambiguous same-part bookmark", true),
            "QUOTE" => new("literal", "refreshes from literal field-code arguments", true),
            "TOC" or "PAGE" or "NUMPAGES" or "SECTIONPAGES" => new("word-layout", "requires Word layout or pagination state", false),
            "DOCPROPERTY" or "DOCVARIABLE" or "AUTHOR" or "TITLE" or "SUBJECT" or "KEYWORDS" => new("document-property", "requires document property state", false),
            "MERGEFIELD" or "MERGEREC" or "MERGESEQ" or "NEXT" or "NEXTIF" or "SKIPIF" => new("mail-merge-data", "requires mail merge data or mail merge state", false),
            "FORMULA" => new("formula", "requires Word formula evaluation", false),
            "IF" => new("conditional", "requires Word conditional field evaluation", false),
            "DATE" or "TIME" or "CREATEDATE" or "SAVEDATE" or "PRINTDATE" => new("date-time", "requires Word date/time evaluation", false),
            "HYPERLINK" or "INCLUDETEXT" or "INCLUDEPICTURE" or "LINK" => new("external", "requires hyperlink or external target state", false),
            _ => new("unsupported", "field type is not modeled for deterministic refresh", false)
        };
    }

    private static void AddFirstFieldOperand(string[] tokens, List<string> dependencies)
    {
        foreach (string token in tokens.Skip(1))
        {
            if (token.StartsWith('\\'))
            {
                continue;
            }

            dependencies.Add(token);
            return;
        }
    }

    private static void ReadHyperlinkFieldDependencies(
        string[] tokens,
        List<string> hyperlinkDependencies,
        List<string> bookmarkDependencies)
    {
        bool expectsAnchor = false;
        for (int i = 1; i < tokens.Length; i++)
        {
            string token = tokens[i];
            if (expectsAnchor)
            {
                bookmarkDependencies.Add(token);
                expectsAnchor = false;
                continue;
            }

            if (string.Equals(token, "\\l", StringComparison.OrdinalIgnoreCase))
            {
                expectsAnchor = true;
                continue;
            }

            if (token.StartsWith('\\'))
            {
                continue;
            }

            hyperlinkDependencies.Add(token);
        }
    }

    private static string[] DistinctNonEmpty(IEnumerable<string> values)
    {
        return values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }

    private static string DetermineFieldSafeEditStatus(string kind, bool complete, bool? isLocked)
    {
        if (!complete)
        {
            return "malformed";
        }

        if (isLocked == true)
        {
            return "locked";
        }

        return kind == "simple" ? "simple-code-result" : "flags-only";
    }

    private static bool? ReadOnOffAttribute(XElement element, string localName)
    {
        // Nullable wrapper around WmlBoolean: a missing attribute reads as unknown
        // (field/history flags are tri-state), not as false.
        string? value = (string?)element.Attribute(OoxmlNs.W + localName);
        return value is null ? null : WmlBoolean.IsTrue(value, valueWhenMissing: false);
    }

    private static string ReadContentControlKind(XElement? properties)
    {
        if (properties is null)
        {
            return "rich-text";
        }

        string? kind = properties.Elements()
            .Select(element => element.Name.LocalName)
            .FirstOrDefault(name => name is "text" or "richText" or "checkBox" or "dropDownList" or "comboBox" or "date" or "picture" or "group" or "repeatingSection" or "repeatingSectionItem");
        return kind switch
        {
            "text" => "plain-text",
            "richText" => "rich-text",
            "checkBox" => "checkbox",
            "dropDownList" => "dropdown-list",
            "comboBox" => "combo-box",
            "date" => "date",
            "picture" => "picture",
            "group" => "group",
            "repeatingSection" => "repeating-section",
            "repeatingSectionItem" => "repeating-section-item",
            _ => "rich-text"
        };
    }

    private static string? ReadSdtProperty(XElement? properties, string localName)
    {
        XElement? element = properties?.Elements().FirstOrDefault(element => element.Name.LocalName == localName);
        return (string?)element?.Attribute(OoxmlNs.W + "val");
    }

    private static string? ReadSdtLock(XElement? properties)
    {
        XElement? element = properties?.Elements().FirstOrDefault(element => element.Name.LocalName == "lock");
        if (element is null)
        {
            return null;
        }

        return (string?)element.Attribute(OoxmlNs.W + "val") ?? "locked";
    }

    private static string? ReadNestedSdtProperty(XElement? properties, string parentLocalName, string localName)
    {
        XElement? element = properties
            ?.Elements()
            .FirstOrDefault(element => element.Name.LocalName == parentLocalName)
            ?.Elements()
            .FirstOrDefault(element => element.Name.LocalName == localName);
        return (string?)element?.Attribute(OoxmlNs.W + "val");
    }

    private static string? ReadSdtAttribute(XElement? properties, string localName, string attributeLocalName)
    {
        XElement? element = properties?.Elements().FirstOrDefault(element => element.Name.LocalName == localName);
        return (string?)element?.Attribute(OoxmlNs.W + attributeLocalName);
    }

    private static string? ReadPlaceholderDocPart(XElement? properties)
    {
        return (string?)properties
            ?.Element(OoxmlNs.W + "placeholder")
            ?.Element(OoxmlNs.W + "docPart")
            ?.Attribute(OoxmlNs.W + "val");
    }

    private static string ReadContentControlSafeEditStatus(string kind, string? lockValue)
    {
        if (!string.IsNullOrWhiteSpace(lockValue) &&
            !string.Equals(lockValue, "unlocked", StringComparison.Ordinal))
        {
            return "locked";
        }

        return kind switch
        {
            "plain-text" => "plain-text",
            "checkbox" => "checkbox",
            "dropdown-list" => "choice",
            "combo-box" => "choice",
            "date" => "date",
            "rich-text" => "rich-text",
            "picture" => "unsupported-picture",
            "group" => "unsupported-group",
            "repeating-section" => "unsupported-repeating-section",
            "repeating-section-item" => "unsupported-repeating-section",
            _ => "unsupported-rich-text"
        };
    }

    private static string? ReadContentControlSafeEditReason(string kind, string? lockValue)
    {
        if (!string.IsNullOrWhiteSpace(lockValue) &&
            !string.Equals(lockValue, "unlocked", StringComparison.Ordinal))
        {
            return $"locked by w:lock='{lockValue}'";
        }

        return kind switch
        {
            "picture" => "picture controls preserve a picture container; inspect media/image targets for image edits",
            "group" => "group controls protect a container; target an editable child content control",
            "repeating-section" => "repeating-section item edits require subtree cloning and are not modeled",
            "repeating-section-item" => "repeating-section item edits require subtree cloning and are not modeled",
            _ => null
        };
    }

    private static bool? ReadContentControlChecked(XElement? properties)
    {
        XElement? checkBox = properties?.Element(OoxmlNs.W + "checkBox");
        XElement? checkedElement = checkBox?.Element(OoxmlNs.W + "checked");
        if (checkedElement is null)
        {
            return null;
        }

        string? value = (string?)checkedElement.Attribute(OoxmlNs.W + "val");
        return value switch
        {
            "1" or "true" or "on" => true,
            "0" or "false" or "off" => false,
            _ => null
        };
    }

    private static string? ReadContentControlStateSymbol(XElement? properties, string stateLocalName)
    {
        return (string?)properties
            ?.Element(OoxmlNs.W + "checkBox")
            ?.Element(OoxmlNs.W + stateLocalName)
            ?.Attribute(OoxmlNs.W + "val");
    }

    private static IReadOnlyList<DocxContentControlListItemInfo> ReadContentControlListItems(XElement? properties)
    {
        XElement? list = properties?.Element(OoxmlNs.W + "dropDownList") ?? properties?.Element(OoxmlNs.W + "comboBox");
        if (list is null)
        {
            return [];
        }

        return list
            .Elements(OoxmlNs.W + "listItem")
            .Select(item => new DocxContentControlListItemInfo(
                (string?)item.Attribute(OoxmlNs.W + "displayText"),
                (string?)item.Attribute(OoxmlNs.W + "value")))
            .ToArray();
    }
}
