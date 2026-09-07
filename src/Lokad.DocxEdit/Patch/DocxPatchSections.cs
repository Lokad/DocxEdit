using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static IReadOnlyList<DocxDiagnostic> ExecuteSetSectionColumns(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? countText = ReadRequiredField(operation, "count", diagnostics);
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!int.TryParse(countText, out int count) || count is < 1 or > 4)
        {
            return [Diagnostic(DocxSeverity.Error, "E6201", "Section column count must be between 1 and 4.", operation, target)];
        }

        SectionTarget? sectionTarget = ResolveMainSectionTarget(package, target, cancellationToken);
        if (sectionTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateSectionGuards(operation, target, sectionTarget.SectionProperties, diagnostics))
        {
            return diagnostics;
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && sectionTarget.SectionProperties.Elements(OoxmlNs.W + "sectPrChange").Any())
        {
            if (!TryFallbackToDirectEdit(options, operation, target, "section already contains tracked section property revision markup", diagnostics, ref useTrackedChanges))
            {
                return diagnostics;
            }
        }


        if (useTrackedChanges)
        {
            SetSectionPropertiesWithTrackedChange(
                package,
                sectionTarget.SectionProperties,
                properties => SetSectionColumns(properties, count),
                options,
                generatedRevisionIds,
                cancellationToken);
        }
        else
        {
            SetSectionColumns(sectionTarget.SectionProperties, count);
        }

        SaveMainDocument(package, sectionTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetSectionOrientation(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? orientation = ReadRequiredField(operation, "orientation", diagnostics);
        if (orientation is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!DocxOrientationExtensions.TryParseWireValue(orientation, out DocxOrientation parsedOrientation))
        {
            return [Diagnostic(DocxSeverity.Error, "E6202", "Section orientation must be portrait or landscape.", operation, target)];
        }

        SectionTarget? sectionTarget = ResolveMainSectionTarget(package, target, cancellationToken);
        if (sectionTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateSectionGuards(operation, target, sectionTarget.SectionProperties, diagnostics))
        {
            return diagnostics;
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && sectionTarget.SectionProperties.Elements(OoxmlNs.W + "sectPrChange").Any())
        {
            if (!TryFallbackToDirectEdit(options, operation, target, "section already contains tracked section property revision markup", diagnostics, ref useTrackedChanges))
            {
                return diagnostics;
            }
        }


        if (useTrackedChanges)
        {
            SetSectionPropertiesWithTrackedChange(
                package,
                sectionTarget.SectionProperties,
                properties => SetSectionOrientation(properties, parsedOrientation),
                options,
                generatedRevisionIds,
                cancellationToken);
        }
        else
        {
            SetSectionOrientation(sectionTarget.SectionProperties, parsedOrientation);
        }

        SaveMainDocument(package, sectionTarget.Document);
        return diagnostics;
    }
    private static bool ValidateSectionGuards(
        DocxPatchOperation operation,
        string target,
        XElement sectionProperties,
        List<DocxDiagnostic> diagnostics)
    {
        if (!TryReadPositiveIntegerGuard(operation, "expect-columns", target, diagnostics, out int? expectedColumns))
        {
            return false;
        }

        if (expectedColumns is not null)
        {
            int actualColumns = ReadSectionColumnCount(sectionProperties);
            if (actualColumns != expectedColumns)
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedColumns} section column(s), found {actualColumns}.", operation, target));
            }
        }

        if (operation.Fields.TryGetValue("expect-orientation", out string? expectedOrientationText))
        {
            if (!DocxOrientationExtensions.TryParseWireValue(expectedOrientationText, out DocxOrientation expectedOrientation))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'expect-orientation' must be portrait or landscape.", operation, target));
            }
            else
            {
                DocxOrientation actualOrientation = ReadSectionOrientation(sectionProperties);
                if (actualOrientation != expectedOrientation)
                {
                    diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected section orientation '{expectedOrientation.ToWireValue()}', found '{actualOrientation.ToWireValue()}'.", operation, target));
                }
            }
        }

        return diagnostics.Count == 0;
    }

    private static int ReadSectionColumnCount(XElement sectionProperties)
    {
        string? countText = (string?)sectionProperties
            .Element(OoxmlNs.W + "cols")
            ?.Attribute(OoxmlNs.W + "num");
        return int.TryParse(countText, out int count) && count > 0 ? count : 1;
    }

    private static DocxOrientation ReadSectionOrientation(XElement sectionProperties)
    {
        string? orientationText = (string?)sectionProperties
            .Element(OoxmlNs.W + "pgSz")
            ?.Attribute(OoxmlNs.W + "orient");
        return DocxOrientationExtensions.TryParseWireValue(orientationText, out DocxOrientation orientation)
            ? orientation
            : DocxOrientation.Portrait;
    }

    private static bool TryReadPositiveIntegerGuard(
        DocxPatchOperation operation,
        string fieldName,
        string target,
        List<DocxDiagnostic> diagnostics,
        out int? value)
    {
        value = null;
        if (!operation.Fields.TryGetValue(fieldName, out string? text))
        {
            return true;
        }

        if (!int.TryParse(text, out int parsed) || parsed <= 0)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", $"Field '{fieldName}' must be greater than 0.", operation, target));
            return false;
        }

        value = parsed;
        return true;
    }
    private static void SetSectionColumns(XElement sectionProperties, int count)
    {
        XElement? columns = sectionProperties.Element(OoxmlNs.W + "cols");
        if (columns is null)
        {
            columns = new XElement(OoxmlNs.W + "cols");
            sectionProperties.Add(columns);
        }

        columns.SetAttributeValue(OoxmlNs.W + "num", count.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    private static void SetSectionOrientation(XElement sectionProperties, DocxOrientation orientation)
    {
        XElement? pageSize = sectionProperties.Element(OoxmlNs.W + "pgSz");
        if (pageSize is null)
        {
            pageSize = new XElement(OoxmlNs.W + "pgSz");
            sectionProperties.AddFirst(pageSize);
        }

        DocxOrientation currentOrientation = ReadSectionOrientation(sectionProperties);
        if (currentOrientation != orientation &&
            pageSize.Attribute(OoxmlNs.W + "w") is XAttribute width &&
            pageSize.Attribute(OoxmlNs.W + "h") is XAttribute height)
        {
            (width.Value, height.Value) = (height.Value, width.Value);
        }

        pageSize.SetAttributeValue(OoxmlNs.W + "orient", orientation.ToWireValue());
    }
}
