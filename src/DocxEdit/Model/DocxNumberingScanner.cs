using System.Xml.Linq;
using DocxEdit.Ooxml;

namespace DocxEdit.Model;

internal sealed class DocxNumberingCatalog
{
    public static DocxNumberingCatalog Empty { get; } = new(
        new Dictionary<string, string>(StringComparer.Ordinal),
        new Dictionary<(string AbstractId, int Level), NumberingLevelDefinition>(),
        new Dictionary<(string NumberingId, int Level), NumberingLevelOverride>());

    private readonly IReadOnlyDictionary<string, string> _numberingToAbstract;
    private readonly IReadOnlyDictionary<(string AbstractId, int Level), NumberingLevelDefinition> _levels;
    private readonly IReadOnlyDictionary<(string NumberingId, int Level), NumberingLevelOverride> _overrides;

    private DocxNumberingCatalog(
        IReadOnlyDictionary<string, string> numberingToAbstract,
        IReadOnlyDictionary<(string AbstractId, int Level), NumberingLevelDefinition> levels,
        IReadOnlyDictionary<(string NumberingId, int Level), NumberingLevelOverride> overrides)
    {
        _numberingToAbstract = numberingToAbstract;
        _levels = levels;
        _overrides = overrides;
    }

    public static DocxNumberingCatalog Scan(OoxmlPackage package, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (package.MainDocumentPartName is null)
        {
            return Empty;
        }

        OoxmlRelationship? relationship = package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .FirstOrDefault(relationship =>
                !relationship.IsExternal &&
                relationship.Type == OoxmlRelTypes.Numbering &&
                relationship.ResolvedTarget is not null);
        if (relationship?.ResolvedTarget is null)
        {
            return Empty;
        }

        OoxmlPart? numberingPart = package.GetPart(relationship.ResolvedTarget);
        if (numberingPart is null)
        {
            return Empty;
        }

        using Stream stream = numberingPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        var numberingToAbstract = new Dictionary<string, string>(StringComparer.Ordinal);
        var levels = new Dictionary<(string AbstractId, int Level), NumberingLevelDefinition>();
        var overrides = new Dictionary<(string NumberingId, int Level), NumberingLevelOverride>();

        foreach (XElement abstractNumbering in document.Root?.Elements(OoxmlNs.W + "abstractNum") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? abstractId = (string?)abstractNumbering.Attribute(OoxmlNs.W + "abstractNumId");
            if (string.IsNullOrWhiteSpace(abstractId))
            {
                continue;
            }

            foreach (XElement level in abstractNumbering.Elements(OoxmlNs.W + "lvl"))
            {
                if (!TryReadLevelIndex(level, out int levelIndex))
                {
                    continue;
                }

                levels[(abstractId, levelIndex)] = ReadLevelDefinition(level);
            }
        }

        foreach (XElement numbering in document.Root?.Elements(OoxmlNs.W + "num") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? numberingId = (string?)numbering.Attribute(OoxmlNs.W + "numId");
            string? abstractId = (string?)numbering
                .Element(OoxmlNs.W + "abstractNumId")
                ?.Attribute(OoxmlNs.W + "val");
            if (string.IsNullOrWhiteSpace(numberingId) || string.IsNullOrWhiteSpace(abstractId))
            {
                continue;
            }

            numberingToAbstract[numberingId] = abstractId;
            foreach (XElement overrideElement in numbering.Elements(OoxmlNs.W + "lvlOverride"))
            {
                if (!TryReadLevelIndex(overrideElement, out int levelIndex))
                {
                    continue;
                }

                int? startOverride = TryReadInt((string?)overrideElement
                    .Element(OoxmlNs.W + "startOverride")
                    ?.Attribute(OoxmlNs.W + "val"));
                XElement? overrideLevel = overrideElement.Element(OoxmlNs.W + "lvl");
                if (startOverride is not null || overrideLevel is not null)
                {
                    overrides[(numberingId, levelIndex)] = new NumberingLevelOverride(
                        startOverride,
                        overrideLevel is null ? null : ReadLevelDefinition(overrideLevel));
                }
            }
        }

        return new DocxNumberingCatalog(numberingToAbstract, levels, overrides);
    }

    public DocxListInfo Resolve(string numberingId, int level, string source)
    {
        string? abstractId = _numberingToAbstract.TryGetValue(numberingId, out string? resolvedAbstractId)
            ? resolvedAbstractId
            : null;
        NumberingLevelDefinition? definition = abstractId is not null &&
            _levels.TryGetValue((abstractId, level), out NumberingLevelDefinition? resolvedDefinition)
                ? resolvedDefinition
                : null;
        NumberingLevelOverride? levelOverride = _overrides.TryGetValue((numberingId, level), out NumberingLevelOverride? resolvedOverride)
            ? resolvedOverride
            : null;
        definition = levelOverride?.Definition ?? definition;
        return new DocxListInfo(numberingId, level)
        {
            AbstractNumberingId = abstractId,
            Format = definition?.Format,
            LevelText = definition?.LevelText,
            StartValue = levelOverride?.StartOverride ?? definition?.StartValue,
            Suffix = definition?.Suffix,
            IsLegal = definition?.IsLegal ?? false,
            RestartAfterLevel = definition?.RestartAfterLevel,
            ParagraphStyleId = definition?.ParagraphStyleId,
            Source = source
        };
    }

    private static bool TryReadLevelIndex(XElement element, out int level)
    {
        string? levelText = (string?)element.Attribute(OoxmlNs.W + "ilvl");
        return int.TryParse(levelText, out level) && level >= 0;
    }

    private static NumberingLevelDefinition ReadLevelDefinition(XElement level)
    {
        return new NumberingLevelDefinition(
            (string?)level.Element(OoxmlNs.W + "numFmt")?.Attribute(OoxmlNs.W + "val"),
            (string?)level.Element(OoxmlNs.W + "lvlText")?.Attribute(OoxmlNs.W + "val"),
            (string?)level.Element(OoxmlNs.W + "pStyle")?.Attribute(OoxmlNs.W + "val"),
            TryReadInt((string?)level.Element(OoxmlNs.W + "start")?.Attribute(OoxmlNs.W + "val")),
            (string?)level.Element(OoxmlNs.W + "suff")?.Attribute(OoxmlNs.W + "val"),
            level.Element(OoxmlNs.W + "isLgl") is not null,
            TryReadInt((string?)level.Element(OoxmlNs.W + "lvlRestart")?.Attribute(OoxmlNs.W + "val")));
    }

    private static int? TryReadInt(string? value)
    {
        return int.TryParse(value, out int parsed) ? parsed : null;
    }

    private sealed record NumberingLevelDefinition(
        string? Format,
        string? LevelText,
        string? ParagraphStyleId,
        int? StartValue,
        string? Suffix,
        bool IsLegal,
        int? RestartAfterLevel);

    private sealed record NumberingLevelOverride(
        int? StartOverride,
        NumberingLevelDefinition? Definition);
}
