using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxPackageValidator
{
    // Structural range checks reused after patch mutation (PLAN B04): paired
    // bookmark/comment-range markers, complex-field balance, and table grid.
    // Reported codes match what Validate would emit for the saved file; the
    // patch engine surfaces only violations that are new relative to the
    // pre-patch baseline, so pre-existing damage never fails a patch.
    internal static IReadOnlyList<DocxDiagnostic> ValidateRangeStructure(XDocument document, string partName)
    {
        var diagnostics = new List<DocxDiagnostic>();
        ValidatePairedIds(document, OoxmlNs.W + "bookmarkStart", OoxmlNs.W + "bookmarkEnd", "bookmark", partName, diagnostics);
        ValidatePairedIds(document, OoxmlNs.W + "commentRangeStart", OoxmlNs.W + "commentRangeEnd", "comment range", partName, diagnostics);
        ValidateFieldBalance(document, partName, diagnostics);
        ValidateTables(document, partName, diagnostics);
        return diagnostics;
    }
}
