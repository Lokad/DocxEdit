using System.Xml.Linq;

namespace Lokad.DocxEdit.Ooxml;

internal static class ContentControlKinds
{
    public static string Read(XElement? properties)
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
}
