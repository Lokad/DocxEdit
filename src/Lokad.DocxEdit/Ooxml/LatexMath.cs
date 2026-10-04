using System.Xml;
using System.Xml.Linq;

namespace Lokad.DocxEdit.Ooxml;

// A deliberately bounded expression language, not a TeX engine. All output is
// native Office Math; unsupported commands fail rather than becoming plain text.
internal sealed class LatexMath(string source)
{
    private int position;
    private int depth;
    private static readonly XNamespace M = OoxmlNs.M;
    private static readonly IReadOnlyDictionary<string, string> Symbols = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["alpha"] = "α", ["beta"] = "β", ["gamma"] = "γ", ["delta"] = "δ", ["epsilon"] = "ϵ",
        ["varepsilon"] = "ε", ["zeta"] = "ζ", ["eta"] = "η", ["theta"] = "θ", ["vartheta"] = "ϑ",
        ["iota"] = "ι", ["kappa"] = "κ", ["lambda"] = "λ", ["mu"] = "μ", ["nu"] = "ν",
        ["xi"] = "ξ", ["pi"] = "π", ["rho"] = "ρ", ["sigma"] = "σ", ["tau"] = "τ",
        ["upsilon"] = "υ", ["phi"] = "ϕ", ["varphi"] = "φ", ["chi"] = "χ", ["psi"] = "ψ", ["omega"] = "ω",
        ["Gamma"] = "Γ", ["Delta"] = "Δ", ["Theta"] = "Θ", ["Lambda"] = "Λ", ["Xi"] = "Ξ",
        ["Pi"] = "Π", ["Sigma"] = "Σ", ["Upsilon"] = "Υ", ["Phi"] = "Φ", ["Psi"] = "Ψ", ["Omega"] = "Ω",
        ["times"] = "×", ["cdot"] = "⋅", ["pm"] = "±", ["mp"] = "∓", ["div"] = "÷",
        ["le"] = "≤", ["leq"] = "≤", ["ge"] = "≥", ["geq"] = "≥", ["ne"] = "≠", ["neq"] = "≠",
        ["approx"] = "≈", ["equiv"] = "≡", ["infty"] = "∞", ["partial"] = "∂", ["nabla"] = "∇",
        ["in"] = "∈", ["notin"] = "∉", ["subset"] = "⊂", ["subseteq"] = "⊆", ["cup"] = "∪", ["cap"] = "∩",
        ["to"] = "→", ["rightarrow"] = "→", ["leftarrow"] = "←", ["Rightarrow"] = "⇒", ["forall"] = "∀",
        ["exists"] = "∃", ["ldots"] = "…", ["cdots"] = "⋯", ["langle"] = "⟨", ["rangle"] = "⟩",
    };

    internal static XElement Parse(string source)
    {
        if (source.Length > 32_768) throw new FormatException("Equation source exceeds 32768 characters.");
        try { XmlConvert.VerifyXmlChars(source); }
        catch (XmlException) { throw new FormatException("Equation source contains an invalid XML character."); }
        var parser = new LatexMath(source);
        List<XElement> content = parser.Expression();
        if (parser.position != source.Length) throw parser.Error("Unexpected closing delimiter");
        if (content.Count == 0) throw parser.Error("An equation must contain an expression");
        return new XElement(M + "oMath", new XAttribute(XNamespace.Xmlns + "m", M), content);
    }

    private FormatException Error(string message) => new($"{message} at character {position + 1}.");
    private bool Starts(string value) => source.AsSpan(position).StartsWith(value, StringComparison.Ordinal);
    private void Space() { while (position < source.Length && char.IsWhiteSpace(source[position])) position++; }

    private List<XElement> Expression(char stop = '\0', bool matrix = false, bool delimiter = false)
    {
        if (++depth > 64) throw Error("Equation nesting exceeds 64 levels");
        var result = new List<XElement>();
        while (true)
        {
            Space();
            if (position == source.Length || source[position] == stop || source[position] == '}' ||
                (matrix && (source[position] == '&' || Starts(@"\\") || Starts(@"\end"))) ||
                (delimiter && Starts(@"\right"))) break;
            List<XElement> atom = Atom();
            (List<XElement>? sub, List<XElement>? sup) = Scripts();
            if (sub is not null || sup is not null)
            {
                atom = [new XElement(M + (sub is null ? "sSup" : sup is null ? "sSub" : "sSubSup"),
                    new XElement(M + "e", atom),
                    sub is null ? null : new XElement(M + "sub", sub),
                    sup is null ? null : new XElement(M + "sup", sup))];
            }
            result.AddRange(atom);
        }
        depth--;
        return result;
    }

    private List<XElement> Group()
    {
        Space();
        if (position == source.Length || source[position++] != '{') throw Error("Expected a braced argument");
        List<XElement> result = Expression('}');
        if (position == source.Length || source[position++] != '}') throw Error("Unclosed argument");
        if (result.Count == 0) throw Error("Empty arguments are not supported");
        return result;
    }

    private (List<XElement>? Sub, List<XElement>? Sup) Scripts()
    {
        List<XElement>? sub = null, sup = null;
        Space();
        while (position < source.Length && source[position] is '_' or '^')
        {
            bool lower = source[position++] == '_';
            Space();
            if (lower ? sub is not null : sup is not null) throw Error("Duplicate subscript or superscript");
            List<XElement> argument = Atom();
            if (lower) sub = argument; else sup = argument;
            Space();
        }
        return (sub, sup);
    }

    private List<XElement> Atom()
    {
        if (++depth > 64) throw Error("Equation nesting exceeds 64 levels");
        try { return AtomCore(); }
        finally { depth--; }
    }

    private List<XElement> AtomCore()
    {
        Space();
        if (position == source.Length) throw Error("Missing expression");
        if (source[position] == '{') return Group();
        char c = source[position++];
        if (c == '\\') return Command();
        if (c is '}' or '^' or '_' or '&' or '$' or '#' or '%') throw Error("Unexpected special character; use a supported escape");
        string value = c.ToString();
        if (char.IsHighSurrogate(c)) value += source[position++];
        return [Run(value)];
    }

    private List<XElement> Command()
    {
        int start = position;
        while (position < source.Length && char.IsAsciiLetter(source[position])) position++;
        if (start == position)
        {
            if (position == source.Length) throw Error("Incomplete command");
            char escape = source[position++];
            return escape switch
            {
                '{' or '}' or '_' or '%' or '#' or '&' or '$' or '|' => [Run(escape.ToString())],
                ',' or ':' or ';' or ' ' => [Run(" ", normal: true)],
                _ => throw Error("Unsupported escape"),
            };
        }
        string command = source[start..position];
        if (Symbols.TryGetValue(command, out string? symbol)) return [Run(symbol)];
        switch (command)
        {
            case "frac": case "dfrac": case "tfrac":
                return [new XElement(M + "f", new XElement(M + "num", Group()), new XElement(M + "den", Group()))];
            case "sqrt":
                Space();
                List<XElement>? degree = null;
                if (position < source.Length && source[position] == '[')
                {
                    position++;
                    degree = Expression(']');
                    if (position == source.Length || source[position++] != ']' || degree.Count == 0) throw Error("Invalid root degree");
                }
                return [new XElement(M + "rad", new XElement(M + "radPr", Property("degHide", degree is null ? "1" : "0")),
                    new XElement(M + "deg", degree), new XElement(M + "e", Group()))];
            case "sum": case "prod": case "int":
                var (sub, sup) = Scripts();
                return [new XElement(M + "nary",
                    new XElement(M + "naryPr", Property("chr", command == "sum" ? "∑" : command == "prod" ? "∏" : "∫"),
                        Property("limLoc", command == "int" ? "subSup" : "undOvr"),
                        Property("subHide", sub is null ? "1" : "0"), Property("supHide", sup is null ? "1" : "0")),
                    new XElement(M + "sub", sub), new XElement(M + "sup", sup), new XElement(M + "e", Group()))];
            case "text":
                return [Run(LiteralGroup(), normal: true)];
            case "mathrm": case "mathbf": case "mathit":
                List<XElement> styled = Group();
                foreach (XElement run in styled.SelectMany(e => e.DescendantsAndSelf(M + "r")))
                {
                    run.Element(M + "rPr")?.Remove();
                    run.AddFirst(new XElement(M + "rPr", Property("sty", command == "mathbf" ? "b" : command == "mathit" ? "i" : "p")));
                }
                return styled;
            case "sin": case "cos": case "tan": case "log": case "ln": case "exp":
                return [Run(command, normal: true)];
            case "left":
                string left = Delimiter();
                List<XElement> inner = Expression(delimiter: true);
                if (!Starts(@"\right")) throw Error("Missing \\right");
                position += 6;
                return [Delimited(left, Delimiter(), inner)];
            case "begin":
                return [Matrix(LiteralGroup())];
            default:
                throw Error($"Unsupported command '\\{command}'");
        }
    }

    private string LiteralGroup()
    {
        Space();
        if (position == source.Length || source[position++] != '{') throw Error("Expected a literal braced argument");
        var result = new System.Text.StringBuilder();
        while (position < source.Length && source[position] != '}')
        {
            char c = source[position++];
            if (c == '{') throw Error("Nested literal braces are not supported");
            if (c == '\\')
            {
                if (position == source.Length || !"{}\\_%&#$".Contains(source[position])) throw Error("Unsupported literal escape");
                c = source[position++];
            }
            result.Append(c);
        }
        if (position == source.Length) throw Error("Unclosed literal argument");
        position++;
        if (result.Length == 0) throw Error("Empty literal argument");
        return result.ToString();
    }

    private string Delimiter()
    {
        Space();
        if (position == source.Length) throw Error("Missing delimiter");
        foreach (var (name, value) in new[] { (@"\langle", "⟨"), (@"\rangle", "⟩"), (@"\{", "{"), (@"\}", "}"), (@"\|", "‖") })
        {
            if (Starts(name)) { position += name.Length; return value; }
        }
        char c = source[position++];
        return c == '.' ? "" : "()[]|".Contains(c) ? c.ToString() : throw Error("Unsupported delimiter");
    }

    private XElement Matrix(string environment)
    {
        if (environment is not ("matrix" or "pmatrix" or "bmatrix" or "cases")) throw Error("Unsupported matrix environment");
        var rows = new List<XElement>();
        int columns = 0, cells = 0;
        while (true)
        {
            var row = new XElement(M + "mr");
            while (true)
            {
                if (++cells > 1024) throw Error("Matrix exceeds 1024 cells");
                row.Add(new XElement(M + "e", Expression(matrix: true)));
                if (position == source.Length) throw Error("Unclosed matrix");
                if (source[position] != '&') break;
                position++;
            }
            int width = row.Elements().Count();
            if (columns != 0 && columns != width) throw Error("Matrix rows must have the same number of cells");
            columns = width;
            rows.Add(row);
            if (Starts(@"\end"))
            {
                position += 4;
                if (LiteralGroup() != environment) throw Error("Mismatched matrix environment");
                break;
            }
            if (!Starts(@"\\")) throw Error("Expected a matrix row separator");
            position += 2;
        }
        var matrix = new XElement(M + "m", rows);
        return environment switch
        {
            "pmatrix" => Delimited("(", ")", [matrix]),
            "bmatrix" => Delimited("[", "]", [matrix]),
            "cases" => Delimited("{", "", [matrix]),
            _ => matrix,
        };
    }

    private static XElement Delimited(string left, string right, List<XElement> inner) => new(M + "d",
        new XElement(M + "dPr", Property("begChr", left), Property("endChr", right)), new XElement(M + "e", inner));
    private static XElement Property(string name, string value) => new(M + name, new XAttribute(M + "val", value));
    private static XElement Run(string text, bool normal = false) => new(M + "r",
        normal ? new XElement(M + "rPr", new XElement(M + "nor")) : null,
        new XElement(OoxmlNs.W + "rPr", new XElement(OoxmlNs.W + "rFonts",
            new XAttribute(OoxmlNs.W + "ascii", "Cambria Math"), new XAttribute(OoxmlNs.W + "hAnsi", "Cambria Math"))),
        new XElement(M + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
}
