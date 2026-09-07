using System.Globalization;
using System.Text;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Rendering;

internal static class TextRenderers
{
    public static string RenderRead(DocxDocumentModel model, int maxText)
    {
        var builder = new StringBuilder();
        foreach (DocxSectionInfo section in model.Sections)
        {
            builder.Append(section.Id.ToWireValue()).Append(" section columns=").Append(section.Columns).Append(" orientation=").Append(section.Orientation.ToWireValue()).AppendLine();
        }

        foreach (DocxParagraphInfo paragraph in model.Paragraphs)
        {
            string kind = paragraph.HeadingLevel is null ? "paragraph" : $"heading level={paragraph.HeadingLevel}";
            string style = paragraph.StyleId is null
                ? string.Empty
                : $" styleId={XmlValues.EscapeText(paragraph.StyleId)}";
            string list = paragraph.List is null
                ? string.Empty
                : RenderList(paragraph.List);
            builder.Append(paragraph.Id.ToWireValue()).Append(' ').Append(kind).Append(style).Append(list).Append(" text=\"").Append(XmlValues.EscapeText(Truncate(paragraph.Text, maxText))).AppendLine("\"");
        }

        foreach (DocxBookmarkInfo bookmark in model.Bookmarks)
        {
            string ooxmlId = bookmark.OoxmlId is null ? string.Empty : $" ooxml-id={XmlValues.EscapeText(bookmark.OoxmlId)}";
            string start = bookmark.StartTargetId is { } bookmarkStart ? $" start={bookmarkStart.ToWireValue()}" : " start=unknown";
            string end = bookmark.EndTargetId is { } bookmarkEnd ? $" end={bookmarkEnd.ToWireValue()}" : " end=unknown";
            string duplicateName = bookmark.IsNameDuplicate ? $" name-duplicate=true duplicate-name-bookmark-ids=\"{XmlValues.EscapeText(string.Join(",", bookmark.DuplicateNameBookmarkIds))}\"" : string.Empty;
            builder.Append(bookmark.Id.ToWireValue())
                .Append(" bookmark name=\"")
                .Append(XmlValues.EscapeText(bookmark.Name))
                .Append('"')
                .Append(ooxmlId)
                .Append(" story=\"")
                .Append(XmlValues.EscapeText(bookmark.Story))
                .Append("\" part=")
                .Append(bookmark.PartName)
                .Append(start)
                .Append(end)
                .Append(" complete=")
                .Append(bookmark.IsComplete)
                .Append(duplicateName)
                .AppendLine();
        }

        foreach (DocxContentControlInfo control in model.ContentControls)
        {
            string target = control.TargetId is { } readControlTarget ? $" target={readControlTarget.ToWireValue()}" : " target=unknown";
            string ooxmlId = control.OoxmlId is null ? string.Empty : $" ooxml-id={XmlValues.EscapeText(control.OoxmlId)}";
            string tag = control.Tag is null ? string.Empty : $" tag=\"{XmlValues.EscapeText(control.Tag)}\"";
            string alias = control.Alias is null ? string.Empty : $" alias=\"{XmlValues.EscapeText(control.Alias)}\"";
            string placeholder = control.PlaceholderDocPart is null ? string.Empty : $" placeholder-doc-part=\"{XmlValues.EscapeText(control.PlaceholderDocPart)}\"";
            string showingPlaceholder = control.IsShowingPlaceholderText ? " showing-placeholder=true" : string.Empty;
            string dataBindingXPath = control.DataBindingXPath is null ? string.Empty : $" data-binding-xpath=\"{XmlValues.EscapeText(control.DataBindingXPath)}\"";
            string dataBindingStore = control.DataBindingStoreItemId is null ? string.Empty : $" data-binding-store-item-id=\"{XmlValues.EscapeText(control.DataBindingStoreItemId)}\"";
            string dataBindingPrefixes = control.DataBindingPrefixMappings is null ? string.Empty : $" data-binding-prefixes=\"{XmlValues.EscapeText(control.DataBindingPrefixMappings)}\"";
            string repeatingSectionTitle = control.RepeatingSectionTitle is null ? string.Empty : $" repeating-section-title=\"{XmlValues.EscapeText(control.RepeatingSectionTitle)}\"";
            string repeatingSectionItems = control.RepeatingSectionItemCount is null ? string.Empty : $" repeating-section-items={control.RepeatingSectionItemCount.Value}";
            string parentControl = control.ParentContentControlId is { } readParentControlId ? $" parent-control={readParentControlId.ToWireValue()}" : string.Empty;
            string childControls = control.ChildContentControlIds.Count == 0 ? string.Empty : $" child-controls=\"{XmlValues.EscapeText(string.Join(",", control.ChildContentControlIds))}\"";
            string safeEdit = $" safe-edit={XmlValues.EscapeText(control.SafeEditStatus)}";
            string safeEditReason = control.SafeEditReason is null ? string.Empty : $" safe-edit-reason=\"{XmlValues.EscapeText(control.SafeEditReason)}\"";
            string tagDuplicate = control.IsTagDuplicate ? $" tag-duplicate=true duplicate-tag-control-ids=\"{XmlValues.EscapeText(string.Join(",", control.DuplicateTagControlIds))}\"" : string.Empty;
            string aliasDuplicate = control.IsAliasDuplicate ? $" alias-duplicate=true duplicate-alias-control-ids=\"{XmlValues.EscapeText(string.Join(",", control.DuplicateAliasControlIds))}\"" : string.Empty;
            string locked = control.Lock is null ? string.Empty : $" lock={XmlValues.EscapeText(control.Lock)}";
            string checkedValue = control.Checked is null ? string.Empty : $" checked={control.Checked.Value.ToString().ToLowerInvariant()}";
            string checkedSymbol = control.CheckedSymbol is null ? string.Empty : $" checked-symbol=\"{XmlValues.EscapeText(control.CheckedSymbol)}\"";
            string uncheckedSymbol = control.UncheckedSymbol is null ? string.Empty : $" unchecked-symbol=\"{XmlValues.EscapeText(control.UncheckedSymbol)}\"";
            string listItems = control.ListItems.Count == 0 ? string.Empty : $" list-items={control.ListItems.Count}";
            string dateFormat = control.DateFormat is null ? string.Empty : $" date-format=\"{XmlValues.EscapeText(control.DateFormat)}\"";
            string dateLanguage = control.DateLanguage is null ? string.Empty : $" date-language={XmlValues.EscapeText(control.DateLanguage)}";
            string dateCalendar = control.DateCalendar is null ? string.Empty : $" date-calendar={XmlValues.EscapeText(control.DateCalendar)}";
            string dateValue = control.DateValue is null ? string.Empty : $" date-value=\"{XmlValues.EscapeText(control.DateValue)}\"";
            builder.Append(control.Id.ToWireValue())
                .Append(" content-control kind=")
                .Append(XmlValues.EscapeText(control.Kind))
                .Append(" story=\"")
                .Append(XmlValues.EscapeText(control.Story))
                .Append("\" part=")
                .Append(control.PartName)
                .Append(target)
                .Append(ooxmlId)
                .Append(tag)
                .Append(alias)
                .Append(placeholder)
                .Append(showingPlaceholder)
                .Append(dataBindingXPath)
                .Append(dataBindingStore)
                .Append(dataBindingPrefixes)
                .Append(repeatingSectionTitle)
                .Append(repeatingSectionItems)
                .Append(parentControl)
                .Append(childControls)
                .Append(safeEdit)
                .Append(safeEditReason)
                .Append(tagDuplicate)
                .Append(aliasDuplicate)
                .Append(locked)
                .Append(checkedValue)
                .Append(checkedSymbol)
                .Append(uncheckedSymbol)
                .Append(listItems)
                .Append(dateFormat)
                .Append(dateLanguage)
                .Append(dateCalendar)
                .Append(dateValue)
                .Append(" text-length=")
                .Append(control.TextLength)
                .AppendLine();
        }

        foreach (DocxFieldInfo field in model.Fields)
        {
            string target = field.TargetId is { } readFieldTarget ? $" target={readFieldTarget.ToWireValue()}" : " target=unknown";
            string fieldType = field.FieldType is null ? string.Empty : $" type={XmlValues.EscapeText(field.FieldType)}";
            string arguments = field.Arguments.Count == 0 ? string.Empty : $" arguments=\"{XmlValues.EscapeText(string.Join(",", field.Arguments))}\"";
            string switches = field.Switches.Count == 0 ? string.Empty : $" switches=\"{XmlValues.EscapeText(string.Join(",", field.Switches))}\"";
            string bookmarkDependencies = field.BookmarkDependencies.Count == 0 ? string.Empty : $" bookmark-dependencies=\"{XmlValues.EscapeText(string.Join(",", field.BookmarkDependencies))}\"";
            string hyperlinkDependencies = field.HyperlinkDependencies.Count == 0 ? string.Empty : $" hyperlink-dependencies=\"{XmlValues.EscapeText(string.Join(",", field.HyperlinkDependencies))}\"";
            string refreshReason = field.RefreshReason is null ? string.Empty : $" refresh-reason=\"{XmlValues.EscapeText(field.RefreshReason)}\"";
            string safeEdit = $" safe-edit={XmlValues.EscapeText(field.SafeEditStatus)}";
            string dirty = field.IsDirty is null ? string.Empty : $" dirty={field.IsDirty}";
            string locked = field.IsLocked is null ? string.Empty : $" locked={field.IsLocked}";
            builder.Append(field.Id.ToWireValue())
                .Append(" field kind=")
                .Append(XmlValues.EscapeText(field.Kind))
                .Append(fieldType)
                .Append(" story=\"")
                .Append(XmlValues.EscapeText(field.Story))
                .Append("\" part=")
                .Append(field.PartName)
                .Append(target)
                .Append(" code=\"")
                .Append(XmlValues.EscapeText(field.Code))
                .Append("\" cached-result=\"")
                .Append(XmlValues.EscapeText(field.CachedResultText))
                .Append("\" result-text-length=")
                .Append(field.ResultTextLength)
                .Append(" nesting-depth=")
                .Append(field.NestingDepth)
                .Append(arguments)
                .Append(switches)
                .Append(bookmarkDependencies)
                .Append(hyperlinkDependencies)
                .Append(" refresh-policy=")
                .Append(field.RefreshPolicy.ToWireValue())
                .Append(" deterministic-refresh=")
                .Append(field.CanRefreshDeterministically.ToString().ToLowerInvariant())
                .Append(refreshReason)
                .Append(safeEdit)
                .Append(dirty)
                .Append(locked)
                .Append(" complete=")
                .Append(field.IsComplete)
                .AppendLine();
        }

        foreach (DocxHyperlinkInfo hyperlink in model.Hyperlinks)
        {
            string target = hyperlink.TargetId is { } readHyperlinkTarget ? $" target={readHyperlinkTarget.ToWireValue()}" : " target=unknown";
            string relationshipId = hyperlink.RelationshipId is null ? string.Empty : $" relationship-id={XmlValues.EscapeText(hyperlink.RelationshipId)}";
            string relationshipPart = hyperlink.RelationshipPartName is null ? string.Empty : $" relationship-part={hyperlink.RelationshipPartName}";
            string relationshipTargetMode = hyperlink.RelationshipTargetMode is null ? string.Empty : $" target-mode={XmlValues.EscapeText(hyperlink.RelationshipTargetMode)}";
            string uri = hyperlink.Uri is null ? string.Empty : $" uri=\"{XmlValues.EscapeText(hyperlink.Uri)}\"";
            string uriScheme = hyperlink.UriScheme is null ? string.Empty : $" uri-scheme={XmlValues.EscapeText(hyperlink.UriScheme)}";
            string uriValid = hyperlink.IsUriValid is null ? string.Empty : $" uri-valid={hyperlink.IsUriValid.Value.ToString().ToLowerInvariant()}";
            string uriReason = hyperlink.UriValidationReason is null ? string.Empty : $" uri-reason={XmlValues.EscapeText(hyperlink.UriValidationReason)}";
            string anchor = hyperlink.Anchor is null ? string.Empty : $" anchor=\"{XmlValues.EscapeText(hyperlink.Anchor)}\"";
            string anchorMissing = hyperlink.IsAnchorMissing is null ? string.Empty : $" anchor-missing={hyperlink.IsAnchorMissing.Value.ToString().ToLowerInvariant()}";
            string anchorDuplicate = hyperlink.IsAnchorDuplicate is null ? string.Empty : $" anchor-duplicate={hyperlink.IsAnchorDuplicate.Value.ToString().ToLowerInvariant()}";
            string tooltip = hyperlink.Tooltip is null ? string.Empty : $" tooltip=\"{XmlValues.EscapeText(hyperlink.Tooltip)}\"";
            string targetFrame = hyperlink.TargetFrame is null ? string.Empty : $" target-frame=\"{XmlValues.EscapeText(hyperlink.TargetFrame)}\"";
            string history = hyperlink.History is null ? string.Empty : $" history={hyperlink.History.Value.ToString().ToLowerInvariant()}";
            string targetPart = hyperlink.TargetPartName is null ? string.Empty : $" target-part={hyperlink.TargetPartName}";
            builder.Append(hyperlink.Id.ToWireValue())
                .Append(" hyperlink story=\"")
                .Append(XmlValues.EscapeText(hyperlink.Story))
                .Append("\" part=")
                .Append(hyperlink.PartName)
                .Append(target)
                .Append(relationshipId)
                .Append(relationshipPart)
                .Append(relationshipTargetMode)
                .Append(uri)
                .Append(uriScheme)
                .Append(uriValid)
                .Append(uriReason)
                .Append(anchor)
                .Append(anchorMissing)
                .Append(anchorDuplicate)
                .Append(tooltip)
                .Append(targetFrame)
                .Append(history)
                .Append(targetPart)
                .Append(" external=")
                .Append(hyperlink.IsExternal)
                .Append(" broken=")
                .Append(hyperlink.IsBroken)
                .Append(" display-text-length=")
                .Append(hyperlink.DisplayTextLength)
                .AppendLine();
        }

        foreach (DocxTableInfo table in model.Tables)
        {
            builder.Append(RenderTable(table, maxText)).AppendLine();
            foreach (DocxTableRowInfo row in table.Rows)
            {
                string gridBefore = row.GridBefore == 0 ? string.Empty : $" grid-before={row.GridBefore}";
                string gridAfter = row.GridAfter == 0 ? string.Empty : $" grid-after={row.GridAfter}";
                string header = row.IsHeader ? " header=true" : string.Empty;
                string cantSplit = row.CantSplit ? " cant-split=true" : string.Empty;
                builder.Append("  ")
                    .Append(row.Id.ToWireValue())
                    .Append(" row cells=")
                    .Append(row.CellCount)
                    .Append(gridBefore)
                    .Append(gridAfter)
                    .Append(header)
                    .Append(cantSplit)
                    .AppendLine();
            }

            foreach (DocxTableCellInfo cell in table.Cells)
            {
                string columnSpan = cell.ColumnSpan == 1 ? string.Empty : $" column-span={cell.ColumnSpan}";
                string visualColumnEnd = cell.VisualColumnEndIndex <= cell.ColumnIndex ? string.Empty : $" visual-column-end={cell.VisualColumnEndIndex}";
                string mergeGroup = cell.MergeGroupId is { } cellMergeGroup ? $" merge-group={cellMergeGroup.ToWireValue()}" : string.Empty;
        string verticalMerge = cell.VerticalMerge is { } cellMerge ? $" vertical-merge={cellMerge.ToWireValue()}" : string.Empty;
                string verticalMergeRoot = cell.VerticalMergeRootCellId is { } cellMergeRoot ? $" vertical-merge-root={cellMergeRoot.ToWireValue()}" : string.Empty;
                string nestedTable = cell.HasNestedTable ? " nested-table=true" : string.Empty;
                string physicalColumn = cell.PhysicalColumnIndex == 0 ? string.Empty : $" physical-column={cell.PhysicalColumnIndex}";
                builder.Append("  ")
                    .Append(cell.Id.ToWireValue())
                    .Append(physicalColumn)
                    .Append(columnSpan)
                    .Append(visualColumnEnd)
                    .Append(mergeGroup)
                    .Append(verticalMerge)
                    .Append(verticalMergeRoot)
                    .Append(nestedTable)
                    .Append(" text=\"")
                    .Append(XmlValues.EscapeText(Truncate(cell.Text, maxText)))
                    .AppendLine("\"");
            }
        }

        foreach (DocxImageInfo image in model.Images)
        {
            builder.Append(RenderImage(image, maxText)).AppendLine();
        }

        return builder.ToString();
    }

    public static IReadOnlyList<DocxOutlineItem> BuildOutline(DocxDocumentModel model, int maxText)
    {
        var items = new List<DocxOutlineItem>();
        foreach (DocxParagraphInfo paragraph in model.Paragraphs.Where(paragraph => paragraph.HeadingLevel is not null))
        {
            items.Add(new DocxOutlineItem
            {
                TargetId = paragraph.Id.ToWireValue(),
                Kind = "heading",
                Text = Truncate(paragraph.Text, maxText),
                HeadingLevel = paragraph.HeadingLevel,
                List = paragraph.List
            });
        }

        foreach (DocxTableInfo table in model.Tables)
        {
            items.Add(new DocxOutlineItem
            {
                TargetId = table.Id.ToWireValue(),
                Kind = "table",
                RowCount = table.RowCount,
                Columns = table.ColumnCount,
                StyleId = table.StyleId,
                Caption = table.Caption is null ? null : Truncate(table.Caption, maxText),
                Description = table.Description is null ? null : Truncate(table.Description, maxText),
                GridColumnCount = table.GridColumnCount,
                HasHeaderRow = table.HasHeaderRow,
                HasMergedCells = table.HasMergedCells,
                HasNestedTables = table.HasNestedTables
            });
        }

        foreach (DocxSectionInfo section in model.Sections)
        {
            items.Add(new DocxOutlineItem
            {
                TargetId = section.Id.ToWireValue(),
                Kind = "section",
                Columns = section.Columns,
                Orientation = section.Orientation.ToWireValue()
            });
        }

        foreach (DocxImageInfo image in model.Images)
        {
            items.Add(new DocxOutlineItem
            {
                TargetId = image.Id.ToWireValue(),
                Kind = "image",
                LayoutKind = image.LayoutKind,
                ContainingTargetId = image.ContainingTargetId is null ? null : image.ContainingTargetId.Value.ToWireValue(),
                PartName = image.PartName
            });
        }

        foreach (DocxBookmarkInfo bookmark in model.Bookmarks)
        {
            items.Add(new DocxOutlineItem
            {
                TargetId = bookmark.Id.ToWireValue(),
                Kind = "bookmark",
                Name = bookmark.Name,
                StartTargetId = bookmark.StartTargetId is null ? null : bookmark.StartTargetId.Value.ToWireValue(),
                EndTargetId = bookmark.EndTargetId is null ? null : bookmark.EndTargetId.Value.ToWireValue(),
                IsNameDuplicate = bookmark.IsNameDuplicate,
                DuplicateNameBookmarkIds = bookmark.DuplicateNameBookmarkIds
            });
        }

        foreach (DocxContentControlInfo control in model.ContentControls)
        {
            items.Add(new DocxOutlineItem
            {
                TargetId = control.Id.ToWireValue(),
                Kind = "content-control",
                ControlKind = control.Kind,
                ControlTargetId = control.TargetId is null ? null : control.TargetId.Value.ToWireValue(),
                Tag = control.Tag,
                Alias = control.Alias,
                ParentControlId = control.ParentContentControlId is null ? null : control.ParentContentControlId.Value.ToWireValue(),
                ChildControlIds = control.ChildContentControlIds,
                SafeEditStatus = control.SafeEditStatus,
                SafeEditReason = control.SafeEditReason,
                IsTagDuplicate = control.IsTagDuplicate,
                DuplicateTagControlIds = control.DuplicateTagControlIds,
                IsAliasDuplicate = control.IsAliasDuplicate,
                DuplicateAliasControlIds = control.DuplicateAliasControlIds,
                Checked = control.Checked,
                ListItemCount = control.ListItems.Count == 0 ? null : control.ListItems.Count,
            });
        }

        foreach (DocxFieldInfo field in model.Fields)
        {
            items.Add(new DocxOutlineItem
            {
                TargetId = field.Id.ToWireValue(),
                Kind = "field",
                FieldKind = field.Kind,
                FieldType = field.FieldType,
                FieldTargetId = field.TargetId is null ? null : field.TargetId.Value.ToWireValue(),
                Code = field.Code,
                NestingDepth = field.NestingDepth,
                RefreshPolicy = field.RefreshPolicy.ToWireValue(),
                DeterministicRefresh = field.CanRefreshDeterministically,
                SafeEditStatus = field.SafeEditStatus
            });
        }

        foreach (DocxHyperlinkInfo hyperlink in model.Hyperlinks)
        {
            items.Add(new DocxOutlineItem
            {
                TargetId = hyperlink.Id.ToWireValue(),
                Kind = "hyperlink",
                HyperlinkTarget = hyperlink.TargetId is null ? null : hyperlink.TargetId.Value.ToWireValue(),
                Destination = hyperlink.Uri ?? hyperlink.Anchor ?? hyperlink.TargetPartName,
                IsBroken = hyperlink.IsBroken
            });
        }

        return items;
    }

    public static string FormatOutlineItem(DocxOutlineItem item)
    {
        if (item.Kind == "heading")
        {
            string list = item.List is null ? string.Empty : RenderList(item.List);
            return $"{item.TargetId} heading level={item.HeadingLevel}{list} text=\"{XmlValues.EscapeText(item.Text)}\"";
        }

        if (item.Kind == "table")
        {
            return FormatOutlineTable(item);
        }

        if (item.Kind == "section")
        {
            return $"{item.TargetId} section columns={item.Columns} orientation={item.Orientation}";
        }

        if (item.Kind == "image")
        {
            string target = item.ContainingTargetId is null ? "target=unknown" : $"target={item.ContainingTargetId}";
            return $"{item.TargetId} image layout={XmlValues.EscapeText(item.LayoutKind)} {target} part={item.PartName}";
        }

        if (item.Kind == "bookmark")
        {
            string start = item.StartTargetId is null ? "unknown" : item.StartTargetId;
            string end = item.EndTargetId is null ? "unknown" : item.EndTargetId;
            string duplicateName = item.IsNameDuplicate ? $" name-duplicate=true duplicate-name-bookmark-ids=\"{XmlValues.EscapeText(string.Join(",", item.DuplicateNameBookmarkIds))}\"" : string.Empty;
            return $"{item.TargetId} bookmark name=\"{XmlValues.EscapeText(item.Name)}\" start={start} end={end}{duplicateName}";
        }

        if (item.Kind == "content-control")
        {
            string target = item.ControlTargetId is null ? "unknown" : item.ControlTargetId;
            string tag = item.Tag is null ? string.Empty : $" tag=\"{XmlValues.EscapeText(item.Tag)}\"";
            string alias = item.Alias is null ? string.Empty : $" alias=\"{XmlValues.EscapeText(item.Alias)}\"";
            string parentControl = item.ParentControlId is null ? string.Empty : $" parent-control={item.ParentControlId}";
            string childControls = item.ChildControlIds.Count == 0 ? string.Empty : $" child-controls=\"{XmlValues.EscapeText(string.Join(",", item.ChildControlIds))}\"";
            string safeEdit = $" safe-edit={XmlValues.EscapeText(item.SafeEditStatus)}";
            string safeEditReason = item.SafeEditReason is null ? string.Empty : $" safe-edit-reason=\"{XmlValues.EscapeText(item.SafeEditReason)}\"";
            string tagDuplicate = item.IsTagDuplicate ? $" tag-duplicate=true duplicate-tag-control-ids=\"{XmlValues.EscapeText(string.Join(",", item.DuplicateTagControlIds))}\"" : string.Empty;
            string aliasDuplicate = item.IsAliasDuplicate ? $" alias-duplicate=true duplicate-alias-control-ids=\"{XmlValues.EscapeText(string.Join(",", item.DuplicateAliasControlIds))}\"" : string.Empty;
            string checkedValue = item.Checked is null ? string.Empty : $" checked={item.Checked.Value.ToString().ToLowerInvariant()}";
            string listItems = item.ListItemCount is null || item.ListItemCount == 0 ? string.Empty : $" list-items={item.ListItemCount}";
            return $"{item.TargetId} content-control kind={XmlValues.EscapeText(item.ControlKind)} target={target}{tag}{alias}{parentControl}{childControls}{safeEdit}{safeEditReason}{tagDuplicate}{aliasDuplicate}{checkedValue}{listItems}";
        }

        if (item.Kind == "field")
        {
            string target = item.FieldTargetId is null ? "unknown" : item.FieldTargetId;
            string fieldType = item.FieldType is null ? string.Empty : $" type={XmlValues.EscapeText(item.FieldType)}";
            string safeEdit = $" safe-edit={XmlValues.EscapeText(item.SafeEditStatus)}";
            return $"{item.TargetId} field kind={XmlValues.EscapeText(item.FieldKind)}" + fieldType + $" target={target} code=\"{XmlValues.EscapeText(item.Code)}\" nesting-depth={item.NestingDepth} refresh-policy={item.RefreshPolicy} deterministic-refresh={item.DeterministicRefresh.ToString().ToLowerInvariant()}{safeEdit}";
        }

        if (item.Kind == "hyperlink")
        {
            string target = item.HyperlinkTarget is null ? "unknown" : item.HyperlinkTarget;
            string destination = item.Destination ?? "unknown";
            return $"{item.TargetId} hyperlink target={target} destination=\"{XmlValues.EscapeText(destination)}\" broken={item.IsBroken}";
        }

        return string.Empty;
    }

    private static string FormatOutlineTable(DocxOutlineItem item)
    {
        string style = item.StyleId is null ? string.Empty : $" styleId={XmlValues.EscapeText(item.StyleId)}";
        string caption = item.Caption is null ? string.Empty : $" caption=\"{XmlValues.EscapeText(item.Caption)}\"";
        string description = item.Description is null ? string.Empty : $" description=\"{XmlValues.EscapeText(item.Description)}\"";
        string grid = item.GridColumnCount is null ? string.Empty : $" grid-columns={item.GridColumnCount}";
        string header = item.HasHeaderRow ? " header-row=true" : string.Empty;
        string merged = item.HasMergedCells ? " merged=true" : string.Empty;
        string nested = item.HasNestedTables ? " nested-table=true" : string.Empty;
        return $"{item.TargetId} table rows={item.RowCount} columns={item.Columns}{style}{caption}{description}{grid}{header}{merged}{nested}";
    }

    public static IReadOnlyList<DocxFindMatch> Find(DocxDocumentModel model, string query, int maxText)
    {
        var matches = new List<DocxFindMatch>();
        foreach (DocxParagraphInfo paragraph in model.Paragraphs)
        {
            if (paragraph.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
            {
                matches.Add(new DocxFindMatch
                {
                    TargetId = paragraph.Id.ToWireValue(),
                    Kind = "paragraph",
                    Text = Truncate(paragraph.Text, maxText),
                    List = paragraph.List
                });
            }
        }

        foreach (DocxTableInfo table in model.Tables)
        {
            foreach (DocxTableCellInfo cell in table.Cells)
            {
                if (cell.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
                {
                    matches.Add(new DocxFindMatch
                    {
                        TargetId = cell.Id.ToWireValue(),
                        Kind = "cell",
                        ParentId = table.Id.ToWireValue(),
                        Text = Truncate(cell.Text, maxText)
                    });
                }
            }
        }

        return matches;
    }

    public static string FormatFindMatch(DocxFindMatch match)
    {
        string list = match.List is null ? string.Empty : RenderList(match.List);
        return match.Kind == "cell"
            ? $"{match.TargetId} text=\"{XmlValues.EscapeText(match.Text)}\""
            : $"{match.TargetId}{list} text=\"{XmlValues.EscapeText(match.Text)}\"";
    }

    public static string? Dump(DocxDocumentModel model, IReadOnlyList<DocxChangeInfo> changes, string targetId, bool includeRuns, int maxText)
    {
        DocxParagraphInfo? paragraph = model.Paragraphs.FirstOrDefault(paragraph => string.Equals(paragraph.Id.ToWireValue(), targetId, StringComparison.Ordinal));
        if (paragraph is not null)
        {
            if (!includeRuns)
            {
                return Truncate(paragraph.Text, maxText);
            }

            var builder = new StringBuilder();
            builder.Append("text=\"").Append(XmlValues.EscapeText(Truncate(paragraph.Text, maxText))).AppendLine("\"");
            builder.AppendLine("runs:");
            for (int i = 0; i < paragraph.Runs.Count; i++)
            {
                DocxRunInfo run = paragraph.Runs[i];
                string markup = run.MarkupType is null ? string.Empty : $" markup={run.MarkupType}";
                string revisionId = run.RevisionId is null ? string.Empty : $" revision-id={XmlValues.EscapeText(run.RevisionId)}";
                string author = run.Author is null ? string.Empty : $" author=\"{XmlValues.EscapeText(run.Author)}\"";
                string timestamp = run.TimestampUtc is null ? string.Empty : $" timestamp-utc={run.TimestampUtc:O}";
                string commentId = run.CommentId is null ? string.Empty : $" comment-id={XmlValues.EscapeText(run.CommentId)}";
                string hyperlinkRelationshipId = run.HyperlinkRelationshipId is null ? string.Empty : $" hyperlink-relationship-id={XmlValues.EscapeText(run.HyperlinkRelationshipId)}";
                string hyperlinkAnchor = run.HyperlinkAnchor is null ? string.Empty : $" hyperlink-anchor=\"{XmlValues.EscapeText(run.HyperlinkAnchor)}\"";
                builder.Append("  ")
                    .Append(paragraph.Id.ToWireValue())
                    .Append(".R")
                    .Append((i + 1).ToString("0000"))
                    .Append(markup)
                    .Append(revisionId)
                    .Append(author)
                    .Append(timestamp)
                    .Append(commentId)
                    .Append(hyperlinkRelationshipId)
                    .Append(hyperlinkAnchor)
                    .Append(" text=\"")
                    .Append(XmlValues.EscapeText(Truncate(run.Text, maxText)))
                    .AppendLine("\"");
            }

            AppendTargetChanges(builder, changes, targetId);
            return builder.ToString();
        }

        DocxTableInfo? table = model.Tables.FirstOrDefault(table => string.Equals(table.Id.ToWireValue(), targetId, StringComparison.Ordinal));
        if (table is not null)
        {
            var builder = new StringBuilder();
            builder.AppendJoin(Environment.NewLine, table.Cells.Select(cell => $"{cell.Id.ToWireValue()}: {Truncate(cell.Text, maxText)}"));
            AppendTargetChanges(builder, changes, targetId);
            return builder.ToString();
        }

        DocxTableCellInfo? cell = model.Tables
            .SelectMany(table => table.Cells)
            .FirstOrDefault(cell => string.Equals(cell.Id.ToWireValue(), targetId, StringComparison.Ordinal));
        if (cell is not null)
        {
            var builder = new StringBuilder();
            builder.Append(Truncate(cell.Text, maxText));
            AppendTargetChanges(builder, changes, targetId);
            return builder.ToString();
        }

        DocxChangeInfo? comment = FindCommentChange(changes, targetId);
        if (comment is not null)
        {
            return RenderCommentDump(comment);
        }

        return RenderTargetChanges(changes, targetId);
    }

    private static void AppendTargetChanges(StringBuilder builder, IReadOnlyList<DocxChangeInfo> changes, string targetId)
    {
        string? renderedChanges = RenderTargetChanges(changes, targetId);
        if (renderedChanges is null)
        {
            return;
        }

        if (builder.Length != 0 && !builder.ToString().EndsWith(Environment.NewLine, StringComparison.Ordinal))
        {
            builder.AppendLine();
        }

        builder.Append(renderedChanges);
    }

    private static string? RenderTargetChanges(IReadOnlyList<DocxChangeInfo> changes, string targetId)
    {
        DocxChangeInfo[] targetChanges = changes
            .Where(change => string.Equals(change.TargetId, targetId, StringComparison.Ordinal))
            .OrderBy(change => change.Id.ToWireValue(), StringComparer.Ordinal)
            .ToArray();
        if (targetChanges.Length == 0)
        {
            return null;
        }

        var builder = new StringBuilder();
        builder.AppendLine("changes:");
        foreach (DocxChangeInfo change in targetChanges)
        {
            string parent = change.ParentType is null ? string.Empty : $" parent={XmlValues.EscapeText(change.ParentType)}";
            string revisionId = change.RevisionId is null ? string.Empty : $" revision-id={XmlValues.EscapeText(change.RevisionId)}";
            string author = change.Author is null ? string.Empty : $" author=\"{XmlValues.EscapeText(change.Author)}\"";
            string timestamp = change.TimestampUtc is null ? string.Empty : $" timestamp-utc={change.TimestampUtc:O}";
            string operationIndex = change.OperationIndex is null ? string.Empty : $" operation-index={change.OperationIndex}";
            string operationName = change.OperationName is null ? string.Empty : $" operation-name={XmlValues.EscapeText(change.OperationName)}";
            string operationTarget = change.OperationTarget is null ? string.Empty : $" operation-target={change.OperationTarget}";
            builder.Append("  ")
                .Append(change.Id.ToWireValue())
                .Append(" type=")
                .Append(XmlValues.EscapeText(change.Type))
                .Append(parent)
                .Append(revisionId)
                .Append(author)
                .Append(timestamp)
                .Append(operationIndex)
                .Append(operationName)
                .Append(operationTarget)
                .Append(" child-elements=")
                .Append(change.ChildElementCount)
                .AppendLine();
        }

        return builder.ToString();
    }

    public static IReadOnlyList<DocxDumpRunInfo> DumpRuns(DocxDocumentModel model, string targetId, int maxText)
    {
        DocxParagraphInfo? paragraph = model.Paragraphs.FirstOrDefault(paragraph => string.Equals(paragraph.Id.ToWireValue(), targetId, StringComparison.Ordinal));
        if (paragraph is null)
        {
            return [];
        }

        return paragraph.Runs
            .Select((run, index) => new DocxDumpRunInfo
            {
                Id = $"{paragraph.Id.ToWireValue()}.R{index + 1:0000}",
                Text = Truncate(run.Text, maxText),
                MarkupType = run.MarkupType,
                RevisionId = run.RevisionId,
                Author = run.Author,
                TimestampUtc = run.TimestampUtc,
                CommentId = run.CommentId,
                HyperlinkRelationshipId = run.HyperlinkRelationshipId,
                HyperlinkAnchor = run.HyperlinkAnchor
            })
            .ToArray();
    }

    public static IReadOnlyList<DocxContextItem> Context(DocxDocumentModel model, IReadOnlyList<DocxChangeInfo> changes, string targetId, int radius, int maxText)
    {
        radius = Math.Max(0, radius);
        TargetAnnotations annotations = BuildTargetAnnotations(model, changes);

        DocxParagraphInfo? paragraph = model.Paragraphs.FirstOrDefault(paragraph => string.Equals(paragraph.Id.ToWireValue(), targetId, StringComparison.Ordinal));
        if (paragraph is not null)
        {
            DocxParagraphInfo[] storyParagraphs = model.Paragraphs
                .Where(candidate => string.Equals(candidate.Story, paragraph.Story, StringComparison.Ordinal))
                .ToArray();
            int index = Array.FindIndex(storyParagraphs, candidate => string.Equals(candidate.Id.ToWireValue(), targetId, StringComparison.Ordinal));
            return Window(storyParagraphs, index, radius)
                .Select(item => ApplyAnnotations(ToContextItem(item.Value, Relation(item.Offset), maxText), annotations))
                .ToArray();
        }

        DocxTableInfo? table = model.Tables.FirstOrDefault(table => string.Equals(table.Id.ToWireValue(), targetId, StringComparison.Ordinal));
        if (table is not null)
        {
            DocxTableInfo[] storyTables = model.Tables
                .Where(candidate => string.Equals(candidate.Story, table.Story, StringComparison.Ordinal))
                .ToArray();
            int index = Array.FindIndex(storyTables, candidate => string.Equals(candidate.Id.ToWireValue(), targetId, StringComparison.Ordinal));
            return Window(storyTables, index, radius)
                .Select(item => ApplyAnnotations(ToContextItem(item.Value, Relation(item.Offset)), annotations))
                .ToArray();
        }

        foreach (DocxTableInfo candidateTable in model.Tables)
        {
            DocxTableCellInfo? cell = candidateTable.Cells.FirstOrDefault(cell => string.Equals(cell.Id.ToWireValue(), targetId, StringComparison.Ordinal));
            if (cell is not null)
            {
                return CellContext(candidateTable, cell, radius, maxText, annotations);
            }

            DocxTableCellInfo[] rowCells = candidateTable.Cells
                .Where(cell => cell.Id.ToWireValue().StartsWith($"{targetId}.C", StringComparison.Ordinal))
                .OrderBy(cell => cell.ColumnIndex)
                .ToArray();
            if (rowCells.Length > 0)
            {
                return RowContext(candidateTable, targetId, rowCells, radius, maxText, annotations);
            }
        }

        DocxSectionInfo? section = model.Sections.FirstOrDefault(section => string.Equals(section.Id.ToWireValue(), targetId, StringComparison.Ordinal));
        if (section is not null)
        {
            DocxSectionInfo[] storySections = model.Sections
                .Where(candidate => string.Equals(candidate.Story, section.Story, StringComparison.Ordinal))
                .ToArray();
            int index = Array.FindIndex(storySections, candidate => string.Equals(candidate.Id.ToWireValue(), targetId, StringComparison.Ordinal));
            return Window(storySections, index, radius)
                .Select(item => ApplyAnnotations(ToContextItem(item.Value, Relation(item.Offset)), annotations))
                .ToArray();
        }

        DocxImageInfo? image = model.Images.FirstOrDefault(image => string.Equals(image.Id.ToWireValue(), targetId, StringComparison.Ordinal));
        if (image is not null)
        {
            return [ToContextItem(image, "target")];
        }

        DocxChangeInfo? comment = FindCommentChange(changes, targetId);
        return comment is null
            ? []
            : [ApplyAnnotations(ToContextItem(comment, "target"), annotations)];
    }

    public static string RenderContext(IReadOnlyList<DocxContextItem> items)
    {
        var builder = new StringBuilder();
        foreach (DocxContextItem item in items)
        {
            string story = string.IsNullOrWhiteSpace(item.Story) ? string.Empty : $" story=\"{XmlValues.EscapeText(item.Story)}\"";
            string parent = item.ParentId is null ? string.Empty : $" parent={item.ParentId}";
            string heading = item.HeadingLevel is null ? string.Empty : $" heading-level={item.HeadingLevel}";
            string style = item.StyleId is null ? string.Empty : $" styleId={XmlValues.EscapeText(item.StyleId)}";
            string list = item.List is null ? string.Empty : RenderList(item.List);
            string bookmarks = item.BookmarkNames.Count == 0 ? string.Empty : $" bookmark-names=\"{XmlValues.EscapeText(string.Join(",", item.BookmarkNames))}\"";
            string contentControlIds = item.ContentControlIds.Count == 0 ? string.Empty : $" content-controls=\"{XmlValues.EscapeText(string.Join(",", item.ContentControlIds))}\"";
            string contentControlTags = item.ContentControlTags.Count == 0 ? string.Empty : $" content-control-tags=\"{XmlValues.EscapeText(string.Join(",", item.ContentControlTags))}\"";
            string contentControlAliases = item.ContentControlAliases.Count == 0 ? string.Empty : $" content-control-aliases=\"{XmlValues.EscapeText(string.Join(",", item.ContentControlAliases))}\"";
            string fieldIds = item.FieldIds.Count == 0 ? string.Empty : $" fields=\"{XmlValues.EscapeText(string.Join(",", item.FieldIds))}\"";
            string fieldCodes = item.FieldCodes.Count == 0 ? string.Empty : $" field-codes=\"{XmlValues.EscapeText(string.Join(",", item.FieldCodes))}\"";
            string fieldKinds = item.FieldKinds.Count == 0 ? string.Empty : $" field-kinds=\"{XmlValues.EscapeText(string.Join(",", item.FieldKinds))}\"";
            string fieldTypes = item.FieldTypes.Count == 0 ? string.Empty : $" field-types=\"{XmlValues.EscapeText(string.Join(",", item.FieldTypes))}\"";
            string hyperlinkIds = item.HyperlinkIds.Count == 0 ? string.Empty : $" hyperlinks=\"{XmlValues.EscapeText(string.Join(",", item.HyperlinkIds))}\"";
            string hyperlinkTargets = item.HyperlinkTargets.Count == 0 ? string.Empty : $" hyperlink-targets=\"{XmlValues.EscapeText(string.Join(",", item.HyperlinkTargets))}\"";
            string commentIds = item.CommentIds.Count == 0 ? string.Empty : $" comments=\"{XmlValues.EscapeText(string.Join(",", item.CommentIds))}\"";
            string commentBodyIds = item.CommentBodyIds.Count == 0 ? string.Empty : $" comment-bodies=\"{XmlValues.EscapeText(string.Join(",", item.CommentBodyIds))}\"";
            string commentParaIds = item.CommentParaIds.Count == 0 ? string.Empty : $" comment-para-ids=\"{XmlValues.EscapeText(string.Join(",", item.CommentParaIds))}\"";
            string commentParentParaIds = item.CommentParentParaIds.Count == 0 ? string.Empty : $" comment-parent-para-ids=\"{XmlValues.EscapeText(string.Join(",", item.CommentParentParaIds))}\"";
            string commentRootParaIds = item.CommentRootParaIds.Count == 0 ? string.Empty : $" comment-root-para-ids=\"{XmlValues.EscapeText(string.Join(",", item.CommentRootParaIds))}\"";
            string commentDurableIds = item.CommentDurableIds.Count == 0 ? string.Empty : $" comment-durable-ids=\"{XmlValues.EscapeText(string.Join(",", item.CommentDurableIds))}\"";
            string commentReplyIds = item.CommentReplyIds.Count == 0 ? string.Empty : $" comment-reply-ids=\"{XmlValues.EscapeText(string.Join(",", item.CommentReplyIds))}\"";
            string commentResolvedIds = item.CommentResolvedIds.Count == 0 ? string.Empty : $" comment-resolved-ids=\"{XmlValues.EscapeText(string.Join(",", item.CommentResolvedIds))}\"";
            string caption = item.Caption is null ? string.Empty : $" caption=\"{XmlValues.EscapeText(item.Caption)}\"";
            string description = item.Description is null ? string.Empty : $" description=\"{XmlValues.EscapeText(item.Description)}\"";
            string rowCount = item.RowCount is null ? string.Empty : $" rows={item.RowCount}";
            string columnCount = item.ColumnCount is null ? string.Empty : $" columns={item.ColumnCount}";
            string row = item.RowIndex is null ? string.Empty : $" row={item.RowIndex}";
            string column = item.ColumnIndex is null ? string.Empty : $" column={item.ColumnIndex}";
            string columnSpan = item.ColumnSpan is null or 1 ? string.Empty : $" column-span={item.ColumnSpan}";
            string visualColumnEnd = item.VisualColumnEndIndex is null ? string.Empty : $" visual-column-end={item.VisualColumnEndIndex}";
            string mergeGroup = item.MergeGroupId is null ? string.Empty : $" merge-group={XmlValues.EscapeText(item.MergeGroupId)}";
        string verticalMerge = item.VerticalMerge is { } itemMerge ? $" vertical-merge={itemMerge.ToWireValue()}" : string.Empty;
            string verticalMergeRoot = item.VerticalMergeRootCellId is null ? string.Empty : $" vertical-merge-root={XmlValues.EscapeText(item.VerticalMergeRootCellId)}";
            string nestedTable = item.HasNestedTable ? " nested-table=true" : string.Empty;
            string text = item.Kind is "paragraph" or "cell"
                ? $" text=\"{XmlValues.EscapeText(item.Text)}\""
                : string.Empty;
            builder.Append(item.Relation)
                .Append(' ')
                .Append(item.Id)
                .Append(' ')
                .Append(item.Kind)
                .Append(story)
                .Append(parent)
                .Append(heading)
                .Append(style)
                .Append(list)
                .Append(bookmarks)
                .Append(contentControlIds)
                .Append(contentControlTags)
                .Append(contentControlAliases)
                .Append(fieldIds)
                .Append(fieldCodes)
                .Append(fieldKinds)
                .Append(fieldTypes)
                .Append(hyperlinkIds)
                .Append(hyperlinkTargets)
                .Append(commentIds)
                .Append(commentBodyIds)
                .Append(commentParaIds)
                .Append(commentParentParaIds)
                .Append(commentRootParaIds)
                .Append(commentDurableIds)
                .Append(commentReplyIds)
                .Append(commentResolvedIds)
                .Append(caption)
                .Append(description)
                .Append(rowCount)
                .Append(columnCount)
                .Append(row)
                .Append(column)
                .Append(columnSpan)
                .Append(visualColumnEnd)
                .Append(mergeGroup)
                .Append(verticalMerge)
                .Append(verticalMergeRoot)
                .Append(nestedTable)
                .Append(text)
                .AppendLine();
        }

        return builder.ToString();
    }

    public static IReadOnlyList<string> RenderStyles(IReadOnlyList<DocxStyleInfo> styles)
    {
        return styles
            .OrderBy(style => style.Type, StringComparer.Ordinal)
            .ThenBy(style => style.StyleId, StringComparer.Ordinal)
            .Select(style =>
            {
                string defaultText = style.IsDefault ? " default=true" : string.Empty;
                string basedOn = style.BasedOnStyleId is null ? string.Empty : $" based-on={XmlValues.EscapeText(style.BasedOnStyleId)}";
                string next = style.NextStyleId is null ? string.Empty : $" next={XmlValues.EscapeText(style.NextStyleId)}";
                string linked = style.LinkedStyleId is null ? string.Empty : $" linked={XmlValues.EscapeText(style.LinkedStyleId)}";
                string numbering = style.NumberingId is null
                    ? string.Empty
                    : $" numbering numId={XmlValues.EscapeText(style.NumberingId)} level={style.NumberingLevel ?? 0}";
                return $"{style.Type} styleId={style.StyleId} name=\"{XmlValues.EscapeText(style.Name)}\"{defaultText}{basedOn}{next}{linked}{numbering}";
            })
            .ToArray();
    }

    private static string RenderList(DocxListInfo list)
    {
        string abstractId = list.AbstractNumberingId is null ? string.Empty : $" abstractNumId={XmlValues.EscapeText(list.AbstractNumberingId)}";
        string format = list.Format is null ? string.Empty : $" format={XmlValues.EscapeText(list.Format)}";
        string levelText = list.LevelText is null ? string.Empty : $" level-text=\"{XmlValues.EscapeText(list.LevelText)}\"";
        string label = list.LabelText is null ? string.Empty : $" label=\"{XmlValues.EscapeText(list.LabelText)}\"";
        string labelStatus = list.LabelStatus == DocxLabelStatus.Resolved
            ? string.Empty
            : $" label-status={list.LabelStatus.ToWireValue()}";
        string labelWarnings = list.LabelWarnings.Count == 0
            ? string.Empty
            : $" label-warnings=\"{XmlValues.EscapeText(string.Join(",", list.LabelWarnings))}\"";
        string labelComponents = list.LabelComponents.Count == 0
            ? string.Empty
            : $" label-components=\"{XmlValues.EscapeText(string.Join(",", list.LabelComponents.Select(component => $"{component.Level}:{component.Value}:{component.Format}:{component.Text}")))}\"";
        string start = list.StartValue is null ? string.Empty : $" start={list.StartValue}";
        string suffix = list.Suffix is null ? string.Empty : $" suffix={XmlValues.EscapeText(list.Suffix)}";
        string legal = list.IsLegal ? " legal=true" : string.Empty;
        string restart = list.RestartAfterLevel is null ? string.Empty : $" restart-after-level={list.RestartAfterLevel}";
        string paragraphStyle = list.ParagraphStyleId is null ? string.Empty : $" paragraph-style={XmlValues.EscapeText(list.ParagraphStyleId)}";
        string source = list.Source == DocxLabelSource.Direct
            ? string.Empty
            : $" source={list.Source.ToWireValue()}";
        return $" list numId={XmlValues.EscapeText(list.NumberingId)} level={list.Level}{abstractId}{format}{levelText}{paragraphStyle}{source}{label}{labelStatus}{labelWarnings}{labelComponents}{start}{suffix}{legal}{restart}";
    }

    private static string RenderImage(DocxImageInfo image, int maxText)
    {
        string relationshipId = image.RelationshipId is null ? string.Empty : $" relationship-id={XmlValues.EscapeText(image.RelationshipId)}";
        string target = image.ContainingTargetId is null ? " target=unknown" : $" target={image.ContainingTargetId}";
        string size = image.WidthEmu is null || image.HeightEmu is null ? string.Empty : $" size-emu={image.WidthEmu}x{image.HeightEmu}";
        string name = image.Name is null ? string.Empty : $" name=\"{XmlValues.EscapeText(Truncate(image.Name, maxText))}\"";
        string description = image.Description is null ? string.Empty : $" description=\"{XmlValues.EscapeText(Truncate(image.Description, maxText))}\"";
        string title = image.Title is null ? string.Empty : $" title=\"{XmlValues.EscapeText(Truncate(image.Title, maxText))}\"";
        string wrap = image.WrapMode is null ? string.Empty : $" wrap={XmlValues.EscapeText(image.WrapMode)}";
        string behind = image.BehindDoc ? " behind-doc=true" : string.Empty;
        string layoutMetadata = RenderImageLayoutMetadata(image);
        string crop = RenderCrop(image);
        return $"{image.Id.ToWireValue()} image layout={XmlValues.EscapeText(image.LayoutKind)} part={image.PartName} content-type={image.ContentType ?? "unknown"} bytes={image.ByteLength}{relationshipId}{target}{size}{name}{description}{title}{wrap}{behind}{layoutMetadata}{crop}";
    }

    private static string RenderImage(DocxImageInfo image)
    {
        return RenderImage(image, -1);
    }

    private static string RenderTable(DocxTableInfo table)
    {
        return RenderTable(table, -1);
    }

    private static string RenderImageLayoutMetadata(DocxImageInfo image)
    {
        return RenderLong("wrap-dist-top-emu", image.WrapDistanceTopEmu) +
            RenderLong("wrap-dist-bottom-emu", image.WrapDistanceBottomEmu) +
            RenderLong("wrap-dist-left-emu", image.WrapDistanceLeftEmu) +
            RenderLong("wrap-dist-right-emu", image.WrapDistanceRightEmu) +
            RenderLong("relative-height", image.RelativeHeight) +
            RenderBool("allow-overlap", image.AllowOverlap) +
            RenderBool("lock-aspect", image.LockAspectRatio) +
            RenderString("position-h-relative", image.HorizontalPositionRelativeFrom) +
            RenderLong("position-h-offset-emu", image.HorizontalPositionOffsetEmu) +
            RenderString("position-h-align", image.HorizontalPositionAlign) +
            RenderString("position-v-relative", image.VerticalPositionRelativeFrom) +
            RenderLong("position-v-offset-emu", image.VerticalPositionOffsetEmu) +
            RenderString("position-v-align", image.VerticalPositionAlign);
    }

    private static string RenderCrop(DocxImageInfo image)
    {
        return RenderPercent("crop-left-percent", image.CropLeftPercent) +
            RenderPercent("crop-top-percent", image.CropTopPercent) +
            RenderPercent("crop-right-percent", image.CropRightPercent) +
            RenderPercent("crop-bottom-percent", image.CropBottomPercent);
    }

    private static string RenderPercent(string name, decimal? value)
    {
        return value is null
            ? string.Empty
            : $" {name}={value.Value.ToString("0.###", CultureInfo.InvariantCulture)}";
    }

    private static string RenderLong(string name, long? value)
    {
        return value is null ? string.Empty : $" {name}={value}";
    }

    private static string RenderBool(string name, bool? value)
    {
        return value is null ? string.Empty : $" {name}={value.Value.ToString().ToLowerInvariant()}";
    }

    private static string RenderString(string name, string? value)
    {
        return value is null ? string.Empty : $" {name}={XmlValues.EscapeText(value)}";
    }

    private static string RenderTable(DocxTableInfo table, int maxText)
    {
        string style = table.StyleId is null ? string.Empty : $" styleId={XmlValues.EscapeText(table.StyleId)}";
        string caption = table.Caption is null ? string.Empty : $" caption=\"{XmlValues.EscapeText(Truncate(table.Caption, maxText))}\"";
        string description = table.Description is null ? string.Empty : $" description=\"{XmlValues.EscapeText(Truncate(table.Description, maxText))}\"";
        string grid = table.GridColumnCount is null ? string.Empty : $" grid-columns={table.GridColumnCount}";
        string header = table.HasHeaderRow ? " header-row=true" : string.Empty;
        string merged = table.HasMergedCells ? " merged=true" : string.Empty;
        string nested = table.HasNestedTables ? " nested-table=true" : string.Empty;
        return $"{table.Id.ToWireValue()} table rows={table.RowCount} columns={table.ColumnCount}{style}{caption}{description}{grid}{header}{merged}{nested}";
    }

    private static IReadOnlyList<DocxContextItem> CellContext(DocxTableInfo table, DocxTableCellInfo cell, int radius, int maxText, TargetAnnotations annotations)
    {
        var items = new List<DocxContextItem>
        {
            ApplyAnnotations(ToContextItem(table, "parent"), annotations)
        };
        DocxTableCellInfo[] rowCells = table.Cells
            .Where(candidate => candidate.RowIndex == cell.RowIndex)
            .OrderBy(candidate => candidate.ColumnIndex)
            .ToArray();
        int index = Array.FindIndex(rowCells, candidate => candidate.Id.Equals(cell.Id));
        items.AddRange(Window(rowCells, index, radius)
            .Select(item => ApplyAnnotations(ToContextItem(item.Value, Relation(item.Offset), maxText, table.Id.ToWireValue(), table.Story), annotations)));
        return items;
    }

    private static IReadOnlyList<DocxContextItem> RowContext(DocxTableInfo table, string rowId, IReadOnlyList<DocxTableCellInfo> rowCells, int radius, int maxText, TargetAnnotations annotations)
    {
        var items = new List<DocxContextItem>
        {
            ApplyAnnotations(ToContextItem(table, "parent"), annotations),
            new()
            {
                Id = rowId,
                Kind = "row",
                Relation = "target",
                Story = table.Story,
                ParentId = table.Id.ToWireValue(),
                RowIndex = rowCells[0].RowIndex,
                ColumnCount = rowCells.Count
            }
        };

        items.AddRange(rowCells
            .Take(Math.Max(1, radius * 2 + 1))
            .Select(cell => ApplyAnnotations(ToContextItem(cell, "child", maxText, table.Id.ToWireValue(), table.Story), annotations)));
        return items;
    }

    private static IEnumerable<(T Value, int Offset)> Window<T>(IReadOnlyList<T> items, int index, int radius)
    {
        if (index < 0)
        {
            yield break;
        }

        int start = Math.Max(0, index - radius);
        int end = Math.Min(items.Count - 1, index + radius);
        for (int i = start; i <= end; i++)
        {
            yield return (items[i], i - index);
        }
    }

    private static string Relation(int offset)
    {
        return offset < 0 ? "before" : offset > 0 ? "after" : "target";
    }

    private static TargetAnnotations BuildTargetAnnotations(DocxDocumentModel model, IReadOnlyList<DocxChangeInfo> changes)
    {
        var bookmarkNames = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (DocxBookmarkInfo bookmark in model.Bookmarks)
        {
            AddAnnotation(bookmarkNames, bookmark.StartTargetId?.ToWireValue(), bookmark.Name);
            AddAnnotation(bookmarkNames, bookmark.EndTargetId?.ToWireValue(), bookmark.Name);
        }

        var contentControlIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var contentControlTags = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var contentControlAliases = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var fieldIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var fieldCodes = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var fieldKinds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var fieldTypes = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var hyperlinkIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var hyperlinkTargets = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var commentIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var commentBodyIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var commentParaIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var commentParentParaIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var commentRootParaIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var commentDurableIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var commentReplyIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var commentResolvedIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (DocxContentControlInfo control in model.ContentControls)
        {
            AddAnnotation(contentControlIds, control.TargetId?.ToWireValue(), control.Id.ToWireValue());
            AddAnnotation(contentControlTags, control.TargetId?.ToWireValue(), control.Tag);
            AddAnnotation(contentControlAliases, control.TargetId?.ToWireValue(), control.Alias);
        }

        foreach (DocxFieldInfo field in model.Fields)
        {
            AddAnnotation(fieldIds, field.TargetId?.ToWireValue(), field.Id.ToWireValue());
            AddAnnotation(fieldCodes, field.TargetId?.ToWireValue(), field.Code);
            AddAnnotation(fieldKinds, field.TargetId?.ToWireValue(), field.Kind);
            AddAnnotation(fieldTypes, field.TargetId?.ToWireValue(), field.FieldType);
        }

        foreach (DocxHyperlinkInfo hyperlink in model.Hyperlinks)
        {
            AddAnnotation(hyperlinkIds, hyperlink.TargetId?.ToWireValue(), hyperlink.Id.ToWireValue());
            AddAnnotation(hyperlinkTargets, hyperlink.TargetId?.ToWireValue(), hyperlink.Uri ?? hyperlink.Anchor ?? hyperlink.TargetPartName);
        }

        var commentBodyById = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (DocxChangeInfo change in changes)
        {
            if (!string.Equals(change.Type, "comment", StringComparison.Ordinal) ||
                string.IsNullOrWhiteSpace(change.CommentId) ||
                string.IsNullOrWhiteSpace(change.TargetId))
            {
                continue;
            }

            commentBodyById.TryAdd(change.CommentId, change.TargetId);
        }
        IReadOnlyDictionary<string, DocxChangeInfo> commentByParaId = changes
            .Where(change => string.Equals(change.Type, "comment", StringComparison.Ordinal))
            .WithNonBlankKey(change => change.CommentParaId)
            .GroupBy(pair => pair.Key, pair => pair.Item, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        foreach (DocxChangeInfo change in changes)
        {
            if (string.IsNullOrWhiteSpace(change.CommentId))
            {
                continue;
            }

            commentBodyById.TryGetValue(change.CommentId, out string? bodyId);
            IEnumerable<string?> targets = new[] { change.TargetId, change.CommentAnchorTargetId, change.CommentReferenceTargetId };
            if (change.CommentIsReply == true &&
                !string.IsNullOrWhiteSpace(change.CommentParentParaId) &&
                commentByParaId.TryGetValue(change.CommentParentParaId, out DocxChangeInfo? parentComment))
            {
                targets = targets.Concat(new[] { parentComment.TargetId, parentComment.CommentAnchorTargetId, parentComment.CommentReferenceTargetId });
            }

            foreach (string? target in targets)
            {
                AddAnnotation(commentIds, target, change.CommentId);
                AddAnnotation(commentBodyIds, target, bodyId);
                AddAnnotation(commentParaIds, target, change.CommentParaId);
                AddAnnotation(commentParentParaIds, target, change.CommentParentParaId);
                AddAnnotation(commentRootParaIds, target, change.CommentRootParaId);
                AddAnnotation(commentDurableIds, target, change.CommentDurableId);
                if (change.CommentIsReply == true)
                {
                    AddAnnotation(commentReplyIds, target, change.CommentId);
                }

                if (change.CommentResolved == true)
                {
                    AddAnnotation(commentResolvedIds, target, change.CommentId);
                }
            }
        }

        return new TargetAnnotations(
            ToArrayDictionary(bookmarkNames),
            ToArrayDictionary(contentControlIds),
            ToArrayDictionary(contentControlTags),
            ToArrayDictionary(contentControlAliases),
            ToArrayDictionary(fieldIds),
            ToArrayDictionary(fieldCodes),
            ToArrayDictionary(fieldKinds),
            ToArrayDictionary(fieldTypes),
            ToArrayDictionary(hyperlinkIds),
            ToArrayDictionary(hyperlinkTargets),
            ToArrayDictionary(commentIds),
            ToArrayDictionary(commentBodyIds),
            ToArrayDictionary(commentParaIds),
            ToArrayDictionary(commentParentParaIds),
            ToArrayDictionary(commentRootParaIds),
            ToArrayDictionary(commentDurableIds),
            ToArrayDictionary(commentReplyIds),
            ToArrayDictionary(commentResolvedIds));
    }

    private static DocxContextItem ApplyAnnotations(DocxContextItem item, TargetAnnotations annotations)
    {
        return item with
        {
            BookmarkNames = LookupAnnotations(annotations.BookmarkNamesByTarget, item.Id),
            ContentControlIds = LookupAnnotations(annotations.ContentControlIdsByTarget, item.Id),
            ContentControlTags = LookupAnnotations(annotations.ContentControlTagsByTarget, item.Id),
            ContentControlAliases = LookupAnnotations(annotations.ContentControlAliasesByTarget, item.Id),
            FieldIds = LookupAnnotations(annotations.FieldIdsByTarget, item.Id),
            FieldCodes = LookupAnnotations(annotations.FieldCodesByTarget, item.Id),
            FieldKinds = LookupAnnotations(annotations.FieldKindsByTarget, item.Id),
            FieldTypes = LookupAnnotations(annotations.FieldTypesByTarget, item.Id),
            HyperlinkIds = LookupAnnotations(annotations.HyperlinkIdsByTarget, item.Id),
            HyperlinkTargets = LookupAnnotations(annotations.HyperlinkTargetsByTarget, item.Id),
            CommentIds = LookupAnnotations(annotations.CommentIdsByTarget, item.Id),
            CommentBodyIds = LookupAnnotations(annotations.CommentBodyIdsByTarget, item.Id),
            CommentParaIds = LookupAnnotations(annotations.CommentParaIdsByTarget, item.Id),
            CommentParentParaIds = LookupAnnotations(annotations.CommentParentParaIdsByTarget, item.Id),
            CommentRootParaIds = LookupAnnotations(annotations.CommentRootParaIdsByTarget, item.Id),
            CommentDurableIds = LookupAnnotations(annotations.CommentDurableIdsByTarget, item.Id),
            CommentReplyIds = LookupAnnotations(annotations.CommentReplyIdsByTarget, item.Id),
            CommentResolvedIds = LookupAnnotations(annotations.CommentResolvedIdsByTarget, item.Id)
        };
    }

    private static void AddAnnotation(Dictionary<string, List<string>> annotations, string? targetId, string? value)
    {
        if (string.IsNullOrWhiteSpace(targetId) || string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        if (!annotations.TryGetValue(targetId, out List<string>? values))
        {
            values = [];
            annotations[targetId] = values;
        }

        if (!values.Contains(value, StringComparer.Ordinal))
        {
            values.Add(value);
        }
    }

    private static IReadOnlyDictionary<string, string[]> ToArrayDictionary(Dictionary<string, List<string>> annotations)
    {
        return annotations.ToDictionary(
            pair => pair.Key,
            pair => pair.Value.Order(StringComparer.Ordinal).ToArray(),
            StringComparer.Ordinal);
    }

    private static IReadOnlyList<string> LookupAnnotations(IReadOnlyDictionary<string, string[]> annotations, string targetId)
    {
        return annotations.TryGetValue(targetId, out string[]? values) ? values : [];
    }

    private static DocxContextItem ToContextItem(DocxParagraphInfo paragraph, string relation, int maxText)
    {
        return new DocxContextItem
        {
            Id = paragraph.Id.ToWireValue(),
            Kind = "paragraph",
            Relation = relation,
            Story = paragraph.Story,
            Text = Truncate(paragraph.Text, maxText),
            HeadingLevel = paragraph.HeadingLevel,
            StyleId = paragraph.StyleId,
            StyleName = paragraph.StyleName,
            List = paragraph.List
        };
    }

    private static DocxContextItem ToContextItem(DocxTableInfo table, string relation)
    {
        return new DocxContextItem
        {
            Id = table.Id.ToWireValue(),
            Kind = "table",
            Relation = relation,
            Story = table.Story,
            Caption = table.Caption,
            Description = table.Description,
            RowCount = table.RowCount,
            ColumnCount = table.ColumnCount
        };
    }

    private static DocxContextItem ToContextItem(DocxTableCellInfo cell, string relation, int maxText, string parentId, string story)
    {
        return new DocxContextItem
        {
            Id = cell.Id.ToWireValue(),
            Kind = "cell",
            Relation = relation,
            Story = story,
            ParentId = parentId,
            Text = Truncate(cell.Text, maxText),
            RowIndex = cell.RowIndex,
            ColumnIndex = cell.ColumnIndex,
            ColumnSpan = cell.ColumnSpan,
            VisualColumnEndIndex = cell.VisualColumnEndIndex <= cell.ColumnIndex ? null : cell.VisualColumnEndIndex,
            MergeGroupId = cell.MergeGroupId?.ToWireValue(),
            VerticalMerge = cell.VerticalMerge,
            VerticalMergeRootCellId = cell.VerticalMergeRootCellId?.ToWireValue(),
            HasNestedTable = cell.HasNestedTable
        };
    }

    private static DocxContextItem ToContextItem(DocxSectionInfo section, string relation)
    {
        return new DocxContextItem
        {
            Id = section.Id.ToWireValue(),
            Kind = "section",
            Relation = relation,
            Story = section.Story,
            ColumnCount = section.Columns
        };
    }

    private static DocxContextItem ToContextItem(DocxImageInfo image, string relation)
    {
        return new DocxContextItem
        {
            Id = image.Id.ToWireValue(),
            Kind = "image",
            Relation = relation,
            ParentId = image.PartName,
            Text = image.ContentType ?? string.Empty
        };
    }

    private static DocxContextItem ToContextItem(DocxChangeInfo comment, string relation)
    {
        return new DocxContextItem
        {
            Id = comment.TargetId ?? (comment.CommentId is null ? comment.Id.ToWireValue() : $"comment:{comment.CommentId}"),
            Kind = "comment",
            Relation = relation,
            Story = comment.Story,
            ParentId = comment.CommentAnchorTargetId ?? comment.CommentReferenceTargetId,
            CommentIds = string.IsNullOrWhiteSpace(comment.CommentId) ? [] : [comment.CommentId],
            CommentBodyIds = string.IsNullOrWhiteSpace(comment.TargetId) ? [] : [comment.TargetId]
        };
    }

    private static DocxChangeInfo? FindCommentChange(IReadOnlyList<DocxChangeInfo> changes, string targetId)
    {
        if (targetId.StartsWith("comment:", StringComparison.Ordinal))
        {
            string commentId = targetId["comment:".Length..].Trim();
            return changes.FirstOrDefault(change =>
                string.Equals(change.Type, "comment", StringComparison.Ordinal) &&
                string.Equals(change.CommentId, commentId, StringComparison.Ordinal));
        }

        return changes.FirstOrDefault(change =>
            string.Equals(change.Type, "comment", StringComparison.Ordinal) &&
            string.Equals(change.TargetId, targetId, StringComparison.Ordinal));
    }

    private static string RenderCommentDump(DocxChangeInfo comment)
    {
        string commentId = comment.CommentId is null ? string.Empty : $" comment-id={XmlValues.EscapeText(comment.CommentId)}";
        string story = string.IsNullOrWhiteSpace(comment.Story) ? string.Empty : $" story=\"{XmlValues.EscapeText(comment.Story)}\"";
        string part = string.IsNullOrWhiteSpace(comment.PartName) ? string.Empty : $" part={comment.PartName}";
        string anchor = comment.CommentAnchorTargetId is null ? string.Empty : $" anchor-target={comment.CommentAnchorTargetId}";
        string reference = comment.CommentReferenceTargetId is null ? string.Empty : $" reference-target={comment.CommentReferenceTargetId}";
        string author = comment.CommentAuthor is null ? string.Empty : $" comment-author=\"{XmlValues.EscapeText(comment.CommentAuthor)}\"";
        string initials = comment.CommentInitials is null ? string.Empty : $" comment-initials=\"{XmlValues.EscapeText(comment.CommentInitials)}\"";
        string timestamp = comment.CommentTimestampUtc is null ? string.Empty : $" comment-timestamp-utc={comment.CommentTimestampUtc:O}";
        return $"comment{commentId}{story}{part}{anchor}{reference}{author}{initials}{timestamp} text-length={comment.TextLength}";
    }


    /// <summary>Bounds body text to the request maximum, marking longer text with a suffix. Inspection builders apply this when items are built; rendering already-bounded text with the same bound is a no-op.</summary>
    internal static string Truncate(string text, int maxText)
    {
        if (maxText < 0 || text.Length <= maxText)
        {
            return text;
        }

        if (maxText <= 3)
        {
            return text[..maxText];
        }

        return text[..(maxText - 3)] + "...";
    }

    private sealed record TargetAnnotations(
        IReadOnlyDictionary<string, string[]> BookmarkNamesByTarget,
        IReadOnlyDictionary<string, string[]> ContentControlIdsByTarget,
        IReadOnlyDictionary<string, string[]> ContentControlTagsByTarget,
        IReadOnlyDictionary<string, string[]> ContentControlAliasesByTarget,
        IReadOnlyDictionary<string, string[]> FieldIdsByTarget,
        IReadOnlyDictionary<string, string[]> FieldCodesByTarget,
        IReadOnlyDictionary<string, string[]> FieldKindsByTarget,
        IReadOnlyDictionary<string, string[]> FieldTypesByTarget,
        IReadOnlyDictionary<string, string[]> HyperlinkIdsByTarget,
        IReadOnlyDictionary<string, string[]> HyperlinkTargetsByTarget,
        IReadOnlyDictionary<string, string[]> CommentIdsByTarget,
        IReadOnlyDictionary<string, string[]> CommentBodyIdsByTarget,
        IReadOnlyDictionary<string, string[]> CommentParaIdsByTarget,
        IReadOnlyDictionary<string, string[]> CommentParentParaIdsByTarget,
        IReadOnlyDictionary<string, string[]> CommentRootParaIdsByTarget,
        IReadOnlyDictionary<string, string[]> CommentDurableIdsByTarget,
        IReadOnlyDictionary<string, string[]> CommentReplyIdsByTarget,
        IReadOnlyDictionary<string, string[]> CommentResolvedIdsByTarget);
}
