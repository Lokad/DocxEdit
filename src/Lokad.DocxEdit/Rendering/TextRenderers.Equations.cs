using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Rendering;

internal static partial class TextRenderers
{
    private static string RenderEquationLine(DocxEquationInfo equation, int maxText) =>
        $"{equation.Id} equation {(equation.IsDisplay ? "display" : "inline")} target={equation.TargetId} hash={equation.ContentHash} text=\"{XmlValues.EscapeText(Truncate(equation.Text, maxText))}\"";
}
