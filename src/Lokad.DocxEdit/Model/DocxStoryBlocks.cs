using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

// Shared story-block enumeration for inspection and editing (PLAN B01).
// IDs are physical: every top-level w:p / w:tbl / w:sectPr in document order,
// including blocks wrapped in a single block-level w:ins / w:del / w:moveFrom /
// w:moveTo container, owns one ordinal. Text views filter which blocks are
// reported, but they never renumber the survivors: a Final view omits deleted
// blocks (leaving gaps), an Original view omits inserted blocks, and Markup
// reports everything. Patch resolvers use the same physical enumeration, so a
// discovered ID always resolves to the same element instead of sliding onto a
// neighbour. IDs are positional within the current package bytes and shift when
// earlier structural operations insert or delete blocks in the same patch;
// resolve order is the operation order against the evolving package.
internal static class DocxStoryBlocks
{
    internal readonly record struct StoryBlock(XElement Block, XElement? Wrapper);

    public static IEnumerable<StoryBlock> EnumeratePhysicalBlocks(XElement container)
    {
        foreach (XElement child in container.Elements())
        {
            if (IsStoryBlock(child))
            {
                yield return new StoryBlock(child, null);
                continue;
            }

            if (!IsRevisionBlockContainer(child))
            {
                continue;
            }

            foreach (XElement inner in child.Elements().Where(IsStoryBlock))
            {
                yield return new StoryBlock(inner, child);
            }
        }
    }

    public static bool IsStoryBlock(XElement element)
    {
        return element.Name == OoxmlNs.W + "p" ||
            element.Name == OoxmlNs.W + "tbl" ||
            element.Name == OoxmlNs.W + "sectPr";
    }

    public static bool IsRevisionBlockContainer(XElement element)
    {
        return element.Name.Namespace == OoxmlNs.W &&
            element.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo" &&
            element.Elements().Any(IsStoryBlock);
    }

    public static bool IsVisibleInView(XElement? wrapper, DocxTextView view)
    {
        if (wrapper is null)
        {
            return true;
        }

        return view switch
        {
            DocxTextView.Final => wrapper.Name.LocalName is not ("del" or "moveFrom"),
            DocxTextView.Original => wrapper.Name.LocalName is not ("ins" or "moveTo"),
            DocxTextView.Markup => true,
            _ => wrapper.Name.LocalName is not ("del" or "moveFrom"),
        };
    }

    public static XElement? FindParagraphByPhysicalOrdinal(XElement container, int ordinal)
    {
        if (ordinal < 1)
        {
            return null;
        }

        int current = 0;
        foreach (StoryBlock entry in EnumeratePhysicalBlocks(container))
        {
            if (entry.Block.Name != OoxmlNs.W + "p")
            {
                continue;
            }

            current++;
            if (current == ordinal)
            {
                return entry.Block;
            }
        }

        return null;
    }

    public static XElement? FindTableByPhysicalOrdinal(XElement container, int ordinal)
    {
        if (ordinal < 1)
        {
            return null;
        }

        int current = 0;
        foreach (StoryBlock entry in EnumeratePhysicalBlocks(container))
        {
            if (entry.Block.Name != OoxmlNs.W + "tbl")
            {
                continue;
            }

            current++;
            if (current == ordinal)
            {
                return entry.Block;
            }
        }

        return null;
    }

    public static XElement? FindBlockByPhysicalOrdinal(XElement container, XName blockName, int ordinal)
    {
        if (ordinal < 1)
        {
            return null;
        }

        int current = 0;
        foreach (StoryBlock entry in EnumeratePhysicalBlocks(container))
        {
            if (entry.Block.Name != blockName)
            {
                continue;
            }

            current++;
            if (current == ordinal)
            {
                return entry.Block;
            }
        }

        return null;
    }
}
