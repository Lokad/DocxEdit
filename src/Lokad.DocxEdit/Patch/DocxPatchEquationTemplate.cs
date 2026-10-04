using System.Text;
using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static string? BuildEquationTemplate(StoryDocument story, DocxTargetId id, DocxTargetCapabilities capabilities, TrackChangesMode mode)
    {
        XElement? math = story.Document.Descendants(OoxmlNs.M + "oMath").ElementAtOrDefault(id.Primary - 1);
        if (math is null) return null;
        var builder = new StringBuilder();
        AppendTemplateHeader(builder, capabilities, mode, hasActive: false);
        builder.AppendLine("# Native OMML is not decompiled to LaTeX; supply the complete replacement expression.");
        foreach (string operation in new[] { "replace-equation", "delete-equation" })
        {
            AppendOpBlock(builder, capabilities, operation, SupportOf(capabilities, operation), active: false, block =>
            {
                block.Append("op ").AppendLine(operation);
                block.Append("target ").AppendLine(id.ToWireValue());
                block.Append("expect-hash ").AppendLine(EquationXml.Hash(math));
                if (operation == "replace-equation") block.AppendLine("latex x^2");
                block.AppendLine("end");
            });
        }
        return builder.ToString();
    }
}
