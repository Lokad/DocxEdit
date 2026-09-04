using System.Globalization;

namespace Lokad.DocxEdit.Ooxml;

internal static class OoxmlUnits
{
    public const long EmusPerInch = 914400;
    public const long EmusPerCentimeter = 360000;
    public const long EmusPerPoint = 12700;
    public const double ScreenDpi = 96;

    public static long InchesToEmu(double inches)
    {
        return checked((long)Math.Round(inches * EmusPerInch, MidpointRounding.AwayFromZero));
    }

    public static long CentimetersToEmu(double centimeters)
    {
        return checked((long)Math.Round(centimeters * EmusPerCentimeter, MidpointRounding.AwayFromZero));
    }

    public static long PointsToEmu(double points)
    {
        return checked((long)Math.Round(points * EmusPerPoint, MidpointRounding.AwayFromZero));
    }

    public static long PixelsToEmu(double pixels)
    {
        return PixelsToEmu(pixels, ScreenDpi);
    }

    public static long PixelsToEmu(double pixels, double dpi)
    {
        if (dpi <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(dpi), "DPI must be positive.");
        }

        return InchesToEmu(pixels / dpi);
    }

    public static bool TryParseDimension(string text, out long emus)
    {
        emus = 0;
        ArgumentNullException.ThrowIfNull(text);
        string trimmed = text.Trim();
        string[] suffixes = ["emu", "in", "cm", "pt", "px"];
        foreach (string suffix in suffixes)
        {
            if (!trimmed.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string numberText = trimmed[..^suffix.Length];
            if (!double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
            {
                return false;
            }

            emus = suffix.ToLowerInvariant() switch
            {
                "emu" => checked((long)Math.Round(value, MidpointRounding.AwayFromZero)),
                "in" => InchesToEmu(value),
                "cm" => CentimetersToEmu(value),
                "pt" => PointsToEmu(value),
                "px" => PixelsToEmu(value),
                _ => 0
            };
            return emus >= 0;
        }

        return false;
    }
}

internal static class OoxmlIds
{
    public static string AllocateRelationshipId(IEnumerable<string> existingIds)
    {
        var existing = new HashSet<string>(existingIds, StringComparer.Ordinal);
        for (int i = 1; ; i++)
        {
            string candidate = $"rId{i}";
            if (!existing.Contains(candidate))
            {
                return candidate;
            }
        }
    }
}

internal static class OoxmlMediaParts
{
    public static string AllocateImagePartName(IEnumerable<string> existingPartNames, string contentType)
    {
        string extension = contentType switch
        {
            "image/png" => "png",
            "image/jpeg" => "jpeg",
            _ => throw new ArgumentException($"Unsupported image content type '{contentType}'.", nameof(contentType))
        };
        var existing = new HashSet<string>(existingPartNames.Select(OoxmlPath.NormalizePartName), StringComparer.OrdinalIgnoreCase);
        for (int i = 1; ; i++)
        {
            string candidate = $"/word/media/image{i}.{extension}";
            if (!existing.Contains(candidate))
            {
                return candidate;
            }
        }
    }
}
