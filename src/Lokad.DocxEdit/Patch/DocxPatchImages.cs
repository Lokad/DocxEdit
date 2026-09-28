using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceImage(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? asset = ReadRequiredField(operation, "asset", diagnostics);
        bool hasAlt = operation.Fields.ContainsKey("alt");
        string? alt = operation.Fields.GetValueOrDefault("alt");
        if (asset is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported replace-image target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryReadAsset(options.AssetProvider, asset, options.Quotas.MaxSinglePartBytes, cancellationToken, out byte[] bytes, out string? contentType, out DocxDiagnostic? assetDiagnostic, operation, target))
        {
            return [assetDiagnostic];
        }

        if (imageTarget.Part.ContentType is not null &&
            !string.Equals(imageTarget.Part.ContentType, contentType, StringComparison.OrdinalIgnoreCase))
        {
            return [Diagnostic(DocxSeverity.Error, "E5204", $"Replacing image content type '{imageTarget.Part.ContentType}' with '{contentType}' is not supported for existing media part {imageTarget.Part.Name}.", operation, target)];
        }

        XElement? imageContainer = null;
        if (hasAlt &&
            !TryGetImageDrawingContainer(imageTarget, target, operation, out imageContainer, out DocxDiagnostic? altDiagnostic))
        {
            return [altDiagnostic];
        }


        package.ReplacePartBytes(imageTarget.Part.Name, bytes);
        if (hasAlt && imageContainer is not null)
        {
            SetImageAlt(imageContainer, alt, target);
            SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        }

        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertImageAfter(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? asset = ReadRequiredField(operation, "asset", diagnostics);
        if (asset is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!TryReadAsset(options.AssetProvider, asset, options.Quotas.MaxSinglePartBytes, cancellationToken, out byte[] bytes, out string? contentType, out DocxDiagnostic? assetDiagnostic, operation, target))
        {
            return [assetDiagnostic];
        }

        if (!ValidateImageContentTypeGuard(operation, target, contentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryReadImageExtent(operation, bytes, contentType, out long widthEmus, out long heightEmus, out DocxDiagnostic? dimensionDiagnostic))
        {
            return [dimensionDiagnostic];
        }


        string imagePartName = OoxmlMediaParts.AllocateImagePartName(package.Parts.Keys, contentType);
        string relationshipId = OoxmlIds.AllocateRelationshipId(package.GetRelationships(paragraphTarget.PartName, cancellationToken).Select(relationship => relationship.Id));
        package.AddPart(imagePartName, contentType, bytes, cancellationToken);
        package.AddRelationship(paragraphTarget.PartName, relationshipId, OoxmlRelTypes.Image, GetRelativeRelationshipTarget(paragraphTarget.PartName, imagePartName), targetMode: null, cancellationToken);
        int docPrId = AllocateDrawingDocPrId(paragraphTarget.Document);
        XElement imageParagraph = CreateInlineImageParagraph(relationshipId, docPrId, widthEmus, heightEmus, operation.Fields.GetValueOrDefault("alt") ?? string.Empty);
        paragraphTarget.Paragraph.AddAfterSelf(imageParagraph);
        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static bool TryParseImageDimension(
        DocxPatchOperation operation,
        string fieldName,
        string text,
        bool requirePositive,
        out long emus,
        [NotNullWhen(false)] out DocxDiagnostic? diagnostic)
    {
        diagnostic = null;
        if (!OoxmlUnits.TryParseDimension(text, out emus) || (requirePositive && emus <= 0))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5206", $"Invalid image {fieldName} '{text}'.", operation, operation.Fields.GetValueOrDefault("target"));
            return false;
        }

        return true;
    }

    private static bool TryReadImageExtent(
        DocxPatchOperation operation,
        byte[] bytes,
        string contentType,
        out long widthEmus,
        out long heightEmus,
        [NotNullWhen(false)] out DocxDiagnostic? diagnostic)
    {
        widthEmus = OoxmlUnits.InchesToEmu(1);
        heightEmus = widthEmus;
        diagnostic = null;
        bool hasWidth = operation.Fields.TryGetValue("width", out string? width);
        bool hasHeight = operation.Fields.TryGetValue("height", out string? height);
        bool hasPixelSize = TryReadImagePixelSize(bytes, contentType, out int pixelWidth, out int pixelHeight);
        if (width is not null && !TryParseImageDimension(operation, "width", width, requirePositive: false, out widthEmus, out diagnostic))
        {
            return false;
        }

        if (height is not null && !TryParseImageDimension(operation, "height", height, requirePositive: false, out heightEmus, out diagnostic))
        {
            return false;
        }

        if (hasWidth && !hasHeight)
        {
            heightEmus = hasPixelSize && pixelWidth > 0
                ? checked((long)Math.Round(widthEmus * (pixelHeight / (double)pixelWidth), MidpointRounding.AwayFromZero))
                : widthEmus;
        }
        else if (!hasWidth && hasHeight)
        {
            widthEmus = hasPixelSize && pixelHeight > 0
                ? checked((long)Math.Round(heightEmus * (pixelWidth / (double)pixelHeight), MidpointRounding.AwayFromZero))
                : heightEmus;
        }
        else if (!hasWidth && !hasHeight && hasPixelSize)
        {
            widthEmus = OoxmlUnits.PixelsToEmu(pixelWidth);
            heightEmus = OoxmlUnits.PixelsToEmu(pixelHeight);
        }

        return true;
    }
    private static bool TryReadImagePixelSize(byte[] bytes, string contentType, out int width, out int height)
    {
        width = 0;
        height = 0;
        return contentType switch
        {
            "image/png" => TryReadPngPixelSize(bytes, out width, out height),
            "image/jpeg" => TryReadJpegPixelSize(bytes, out width, out height),
            _ => false
        };
    }

    private static bool TryReadPngPixelSize(byte[] bytes, out int width, out int height)
    {
        width = 0;
        height = 0;
        if (bytes.Length < 24 ||
            bytes[0] != 0x89 ||
            bytes[1] != 0x50 ||
            bytes[2] != 0x4E ||
            bytes[3] != 0x47 ||
            bytes[4] != 0x0D ||
            bytes[5] != 0x0A ||
            bytes[6] != 0x1A ||
            bytes[7] != 0x0A ||
            bytes[12] != 0x49 ||
            bytes[13] != 0x48 ||
            bytes[14] != 0x44 ||
            bytes[15] != 0x52)
        {
            return false;
        }

        width = ReadBigEndianInt32(bytes, 16);
        height = ReadBigEndianInt32(bytes, 20);
        return width > 0 && height > 0;
    }

    private static bool TryReadJpegPixelSize(byte[] bytes, out int width, out int height)
    {
        width = 0;
        height = 0;
        if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
        {
            return false;
        }

        int index = 2;
        while (index + 3 < bytes.Length)
        {
            if (bytes[index] != 0xFF)
            {
                index++;
                continue;
            }

            while (index < bytes.Length && bytes[index] == 0xFF)
            {
                index++;
            }

            if (index >= bytes.Length)
            {
                return false;
            }

            byte marker = bytes[index++];
            if (marker is 0xD9 or 0xDA)
            {
                return false;
            }

            if (marker is 0x01 || marker is >= 0xD0 and <= 0xD7)
            {
                continue;
            }

            if (index + 1 >= bytes.Length)
            {
                return false;
            }

            int segmentLength = ReadBigEndianUInt16(bytes, index);
            if (segmentLength < 2 || index + segmentLength > bytes.Length)
            {
                return false;
            }

            if (IsJpegStartOfFrame(marker) && segmentLength >= 7)
            {
                height = ReadBigEndianUInt16(bytes, index + 3);
                width = ReadBigEndianUInt16(bytes, index + 5);
                return width > 0 && height > 0;
            }

            index += segmentLength;
        }

        return false;
    }

    private static bool IsJpegStartOfFrame(byte marker)
    {
        return marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;
    }

    private static int ReadBigEndianInt32(byte[] bytes, int offset)
    {
        return (bytes[offset] << 24) |
            (bytes[offset + 1] << 16) |
            (bytes[offset + 2] << 8) |
            bytes[offset + 3];
    }

    private static int ReadBigEndianUInt16(byte[] bytes, int offset)
    {
        return (bytes[offset] << 8) | bytes[offset + 1];
    }

    private static string GetRelativeRelationshipTarget(string sourcePartName, string targetPartName)
    {
        string[] sourceSegments = OoxmlPath.NormalizePartName(sourcePartName).Split('/', StringSplitOptions.RemoveEmptyEntries);
        string[] targetSegments = OoxmlPath.NormalizePartName(targetPartName).Split('/', StringSplitOptions.RemoveEmptyEntries);
        int sourceDirectoryLength = sourceSegments.Length - 1;
        int common = 0;
        while (common < sourceDirectoryLength &&
            common < targetSegments.Length - 1 &&
            string.Equals(sourceSegments[common], targetSegments[common], StringComparison.OrdinalIgnoreCase))
        {
            common++;
        }

        var relative = new List<string>();
        for (int i = common; i < sourceDirectoryLength; i++)
        {
            relative.Add("..");
        }

        for (int i = common; i < targetSegments.Length; i++)
        {
            relative.Add(targetSegments[i]);
        }

        return string.Join('/', relative);
    }

    private static int AllocateDrawingDocPrId(XDocument document)
    {
        var existing = new HashSet<int>();
        foreach (XElement docPr in document.Descendants(OoxmlNs.Wp + "docPr"))
        {
            if (int.TryParse((string?)docPr.Attribute("id"), out int id) && id > 0)
            {
                existing.Add(id);
            }
        }

        for (int id = 1; ; id++)
        {
            if (!existing.Contains(id))
            {
                return id;
            }
        }
    }

    private static XElement CreateInlineImageParagraph(string relationshipId, int docPrId, long widthEmus, long heightEmus, string alt)
    {
        return new XElement(
            OoxmlNs.W + "p",
            new XElement(
                OoxmlNs.W + "r",
                new XElement(
                    OoxmlNs.W + "drawing",
                    new XElement(
                        OoxmlNs.Wp + "inline",
                        new XElement(OoxmlNs.Wp + "extent", new XAttribute("cx", widthEmus), new XAttribute("cy", heightEmus)),
                        new XElement(OoxmlNs.Wp + "docPr", new XAttribute("id", docPrId), new XAttribute("name", $"Picture {docPrId}"), new XAttribute("descr", alt)),
                        new XElement(
                            OoxmlNs.A + "graphic",
                            new XElement(
                                OoxmlNs.A + "graphicData",
                                new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
                                new XElement(
                                    OoxmlNs.Pic + "pic",
                                    new XElement(
                                        OoxmlNs.Pic + "nvPicPr",
                                        new XElement(OoxmlNs.Pic + "cNvPr", new XAttribute("id", "0"), new XAttribute("name", "Picture")),
                                        new XElement(OoxmlNs.Pic + "cNvPicPr")),
                                    new XElement(
                                        OoxmlNs.Pic + "blipFill",
                                        new XElement(OoxmlNs.A + "blip", new XAttribute(OoxmlNs.R + "embed", relationshipId)),
                                        new XElement(OoxmlNs.A + "stretch", new XElement(OoxmlNs.A + "fillRect"))),
                                    new XElement(
                                        OoxmlNs.Pic + "spPr",
                                        new XElement(
                                            OoxmlNs.A + "xfrm",
                                            new XElement(OoxmlNs.A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                            new XElement(OoxmlNs.A + "ext", new XAttribute("cx", widthEmus), new XAttribute("cy", heightEmus))),
                                        new XElement(OoxmlNs.A + "prstGeom", new XAttribute("prst", "rect"), new XElement(OoxmlNs.A + "avLst"))))))))));
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageAlt(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? alt = ReadRequiredField(operation, "alt", diagnostics);
        string? expectedAlt = operation.Fields.GetValueOrDefault("expect-alt");
        if (alt is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-alt target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic];
        }

        string? currentAlt = (string?)imageContainer.Element(OoxmlNs.Wp + "docPr")?.Attribute("descr");
        if (expectedAlt is not null && !string.Equals(currentAlt, expectedAlt, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected image alt " + "\u0027" + expectedAlt + "\u0027" + ", found " + "\u0027" + (currentAlt ?? "none") + "\u0027" + ".", operation, target, fieldName: "expect-alt")];
        }


        SetImageAlt(imageContainer, alt, target);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageMetadata(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? alt = operation.Fields.GetValueOrDefault("alt");
        string? title = operation.Fields.GetValueOrDefault("title");
        string? name = operation.Fields.GetValueOrDefault("name");
        string? expectedAlt = operation.Fields.GetValueOrDefault("expect-alt");
        string? expectedTitle = operation.Fields.GetValueOrDefault("expect-title");
        string? expectedName = operation.Fields.GetValueOrDefault("expect-name");
        if (alt is null && title is null && name is null)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-image-metadata' requires at least one of 'alt', 'title', or 'name'.", operation, target));
        }

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-metadata target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic];
        }

        XElement? metadataDocPr = imageContainer.Element(OoxmlNs.Wp + "docPr");
        string? currentAlt = (string?)metadataDocPr?.Attribute("descr");
        string? currentTitle = (string?)metadataDocPr?.Attribute("title");
        string? currentMetadataName = (string?)metadataDocPr?.Attribute("name");
        if (expectedAlt is not null && !string.Equals(currentAlt, expectedAlt, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected image alt " + "\u0027" + expectedAlt + "\u0027" + ", found " + "\u0027" + (currentAlt ?? "none") + "\u0027" + ".", operation, target, fieldName: "expect-alt")];
        }

        if (expectedTitle is not null && !string.Equals(currentTitle, expectedTitle, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected image title " + "\u0027" + expectedTitle + "\u0027" + ", found " + "\u0027" + (currentTitle ?? "none") + "\u0027" + ".", operation, target, fieldName: "expect-title")];
        }

        if (expectedName is not null && !string.Equals(currentMetadataName, expectedName, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected image name " + "\u0027" + expectedName + "\u0027" + ", found " + "\u0027" + (currentMetadataName ?? "none") + "\u0027" + ".", operation, target, fieldName: "expect-name")];
        }


        SetImageMetadata(imageContainer, target, alt, title, name);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageSize(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool hasWidth = operation.Fields.ContainsKey("width");
        bool hasHeight = operation.Fields.ContainsKey("height");
        if (!hasWidth && !hasHeight)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-image-size' requires 'width', 'height', or both.", operation, target));
        }

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-size target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic];
        }

        TryReadExistingImageSize(imageContainer, imageTarget.Part, out long currentWidthEmus, out long currentHeightEmus);
        if (!TryReadImageSize(operation, currentWidthEmus, currentHeightEmus, out long widthEmus, out long heightEmus, out diagnostic))
        {
            return [diagnostic];
        }


        SetImageSize(imageContainer, widthEmus, heightEmus);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageWrap(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool hasWrapField =
            operation.Fields.ContainsKey("mode") ||
            operation.Fields.ContainsKey("dist-top") ||
            operation.Fields.ContainsKey("dist-bottom") ||
            operation.Fields.ContainsKey("dist-left") ||
            operation.Fields.ContainsKey("dist-right");
        if (!hasWrapField)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-image-wrap' requires 'mode' or at least one distance field.", operation, target));
        }

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-wrap target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic];
        }

        if (imageContainer.Name != OoxmlNs.Wp + "anchor")
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' is inline; wrap metadata is only editable on anchored images.", operation, target)];
        }

        string? mode = null;
        if (operation.Fields.TryGetValue("mode", out string? modeText) &&
            !TryNormalizeWrapMode(modeText, out mode))
        {
            return [Diagnostic(DocxSeverity.Error, "E5209", $"Unsupported image wrap mode '{modeText}'.", operation, target)];
        }

        if (!TryReadWrapDistanceField(operation, "dist-top", diagnostics, out long? distanceTop) ||
            !TryReadWrapDistanceField(operation, "dist-bottom", diagnostics, out long? distanceBottom) ||
            !TryReadWrapDistanceField(operation, "dist-left", diagnostics, out long? distanceLeft) ||
            !TryReadWrapDistanceField(operation, "dist-right", diagnostics, out long? distanceRight))
        {
            return diagnostics;
        }


        if (mode is not null)
        {
            SetImageWrapMode(imageContainer, mode);
        }

        SetImageWrapDistance(imageContainer, "distT", distanceTop);
        SetImageWrapDistance(imageContainer, "distB", distanceBottom);
        SetImageWrapDistance(imageContainer, "distL", distanceLeft);
        SetImageWrapDistance(imageContainer, "distR", distanceRight);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImagePosition(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool hasPositionField =
            operation.Fields.ContainsKey("horizontal-relative") ||
            operation.Fields.ContainsKey("horizontal-offset") ||
            operation.Fields.ContainsKey("horizontal-align") ||
            operation.Fields.ContainsKey("vertical-relative") ||
            operation.Fields.ContainsKey("vertical-offset") ||
            operation.Fields.ContainsKey("vertical-align");
        if (!hasPositionField)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-image-position' requires at least one position field.", operation, target));
        }

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-position target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic];
        }

        if (imageContainer.Name != OoxmlNs.Wp + "anchor")
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' is inline; position metadata is only editable on anchored images.", operation, target)];
        }

        if (!TryReadImagePositionAxis(operation, "horizontal", diagnostics, out ImagePositionAxis horizontal) ||
            !TryReadImagePositionAxis(operation, "vertical", diagnostics, out ImagePositionAxis vertical))
        {
            return diagnostics;
        }


        SetImagePositionAxis(imageContainer, "positionH", horizontal);
        SetImagePositionAxis(imageContainer, "positionV", vertical);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageCrop(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool hasCropField =
            operation.Fields.ContainsKey("left-percent") ||
            operation.Fields.ContainsKey("top-percent") ||
            operation.Fields.ContainsKey("right-percent") ||
            operation.Fields.ContainsKey("bottom-percent");
        if (!hasCropField)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-image-crop' requires at least one crop field.", operation, target));
        }

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-crop target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        XElement? blipFill = imageTarget.Blip.Ancestors(OoxmlNs.Pic + "blipFill").FirstOrDefault();
        if (blipFill is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' does not have editable DrawingML crop metadata.", operation, target)];
        }

        ImageCrop crop = ReadImageCrop(blipFill.Element(OoxmlNs.A + "srcRect"));
        if (!TryReadCropPercentField(operation, "left-percent", crop.Left, diagnostics, out int left) ||
            !TryReadCropPercentField(operation, "top-percent", crop.Top, diagnostics, out int top) ||
            !TryReadCropPercentField(operation, "right-percent", crop.Right, diagnostics, out int right) ||
            !TryReadCropPercentField(operation, "bottom-percent", crop.Bottom, diagnostics, out int bottom))
        {
            return diagnostics;
        }

        crop = new ImageCrop(left, top, right, bottom);
        if (crop.Left + crop.Right >= 100_000)
        {
            return [Diagnostic(DocxSeverity.Error, "E5208", "Image crop left-percent plus right-percent must be less than 100.", operation, target)];
        }

        if (crop.Top + crop.Bottom >= 100_000)
        {
            return [Diagnostic(DocxSeverity.Error, "E5208", "Image crop top-percent plus bottom-percent must be less than 100.", operation, target)];
        }


        SetImageCrop(blipFill, crop);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static bool TryGetImageDrawingContainer(
        ImageBlipTarget imageTarget,
        string target,
        DocxPatchOperation operation,
        [NotNullWhen(true)] out XElement? container,
        [NotNullWhen(false)] out DocxDiagnostic? diagnostic)
    {
        XElement? drawing = imageTarget.Blip.Ancestors(OoxmlNs.W + "drawing").FirstOrDefault();
        container = drawing?.Descendants(OoxmlNs.Wp + "inline").FirstOrDefault()
            ?? drawing?.Descendants(OoxmlNs.Wp + "anchor").FirstOrDefault();
        if (container is null)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' does not have editable DrawingML properties.", operation, target);
            return false;
        }

        diagnostic = null;
        return true;
    }

    private static void SetImageAlt(XElement container, string? alt, string target)
    {
        SetImageMetadata(container, target, alt, title: null, name: null);
    }

    private static void SetImageMetadata(XElement container, string target, string? alt, string? title, string? name)
    {
        XElement? docPr = container.Element(OoxmlNs.Wp + "docPr");
        if (docPr is null)
        {
            docPr = new XElement(OoxmlNs.Wp + "docPr");
            docPr.SetAttributeValue("id", "1");
            docPr.SetAttributeValue("name", target);
            container.AddFirst(docPr);
        }

        if (alt is not null)
        {
            docPr.SetAttributeValue("descr", alt);
        }

        if (title is not null)
        {
            docPr.SetAttributeValue("title", title);
        }

        if (name is not null)
        {
            docPr.SetAttributeValue("name", name);
        }
    }

    private static bool TryReadExistingImageSize(XElement container, OoxmlPart part, out long widthEmus, out long heightEmus)
    {
        XElement? extent = container.Element(OoxmlNs.Wp + "extent");
        widthEmus = ReadLongAttribute(extent, "cx") ?? 0;
        heightEmus = ReadLongAttribute(extent, "cy") ?? 0;
        if (widthEmus > 0 && heightEmus > 0)
        {
            return true;
        }

        if (part.ContentType is not null &&
            TryReadImagePixelSize(part.Bytes, part.ContentType, out int pixelWidth, out int pixelHeight))
        {
            widthEmus = OoxmlUnits.PixelsToEmu(pixelWidth);
            heightEmus = OoxmlUnits.PixelsToEmu(pixelHeight);
            return true;
        }

        widthEmus = OoxmlUnits.InchesToEmu(1);
        heightEmus = widthEmus;
        return true;
    }

    private static bool TryReadImageSize(
        DocxPatchOperation operation,
        long currentWidthEmus,
        long currentHeightEmus,
        out long widthEmus,
        out long heightEmus,
        [NotNullWhen(false)] out DocxDiagnostic? diagnostic)
    {
        widthEmus = currentWidthEmus;
        heightEmus = currentHeightEmus;
        diagnostic = null;
        bool hasWidth = operation.Fields.TryGetValue("width", out string? width);
        bool hasHeight = operation.Fields.TryGetValue("height", out string? height);
        if (width is not null && !TryParseImageDimension(operation, "width", width, requirePositive: true, out widthEmus, out diagnostic))
        {
            return false;
        }

        if (height is not null && !TryParseImageDimension(operation, "height", height, requirePositive: true, out heightEmus, out diagnostic))
        {
            return false;
        }

        if (hasWidth && !hasHeight)
        {
            heightEmus = checked((long)Math.Round(widthEmus * (currentHeightEmus / (double)currentWidthEmus), MidpointRounding.AwayFromZero));
        }
        else if (!hasWidth && hasHeight)
        {
            widthEmus = checked((long)Math.Round(heightEmus * (currentWidthEmus / (double)currentHeightEmus), MidpointRounding.AwayFromZero));
        }

        return true;
    }

    private static long? ReadLongAttribute(XElement? element, string localName)
    {
        return long.TryParse((string?)element?.Attribute(localName), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out long value)
            ? value
            : null;
    }

    private static void SetImageSize(XElement container, long widthEmus, long heightEmus)
    {
        XElement extent = container.Element(OoxmlNs.Wp + "extent") ?? new XElement(OoxmlNs.Wp + "extent");
        if (extent.Parent is null)
        {
            container.AddFirst(extent);
        }

        extent.SetAttributeValue("cx", widthEmus);
        extent.SetAttributeValue("cy", heightEmus);

        XElement? picture = container.Descendants(OoxmlNs.Pic + "pic").FirstOrDefault();
        if (picture is null)
        {
            return;
        }

        XElement shapeProperties = picture.Element(OoxmlNs.Pic + "spPr") ?? new XElement(OoxmlNs.Pic + "spPr");
        if (shapeProperties.Parent is null)
        {
            picture.Add(shapeProperties);
        }

        XElement transform = shapeProperties.Element(OoxmlNs.A + "xfrm") ?? new XElement(OoxmlNs.A + "xfrm");
        if (transform.Parent is null)
        {
            shapeProperties.AddFirst(transform);
        }

        XElement transformExtent = transform.Element(OoxmlNs.A + "ext") ?? new XElement(OoxmlNs.A + "ext");
        if (transformExtent.Parent is null)
        {
            transform.Add(transformExtent);
        }

        transformExtent.SetAttributeValue("cx", widthEmus);
        transformExtent.SetAttributeValue("cy", heightEmus);
    }

    private static bool TryNormalizeWrapMode(string text, out string? mode)
    {
        mode = text switch
        {
            "none" or "wrapNone" => "wrapNone",
            "square" or "wrapSquare" => "wrapSquare",
            "tight" or "wrapTight" => "wrapTight",
            "through" or "wrapThrough" => "wrapThrough",
            "top-bottom" or "topAndBottom" or "wrapTopAndBottom" => "wrapTopAndBottom",
            _ => null
        };
        return mode is not null;
    }

    private static bool TryReadWrapDistanceField(
        DocxPatchOperation operation,
        string fieldName,
        List<DocxDiagnostic> diagnostics,
        out long? value)
    {
        value = null;
        if (!operation.Fields.TryGetValue(fieldName, out string? text))
        {
            return true;
        }

        if (!OoxmlUnits.TryParseDimension(text, out long emus))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5209", $"Image wrap distance field '{fieldName}' must be a non-negative dimension.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        value = emus;
        return true;
    }

    private static void SetImageWrapMode(XElement anchor, string mode)
    {
        anchor.Elements()
            .Where(element => element.Name.Namespace == OoxmlNs.Wp && element.Name.LocalName.StartsWith("wrap", StringComparison.Ordinal))
            .Remove();
        XElement wrap = new(OoxmlNs.Wp + mode);
        XElement? insertAfter = anchor.Element(OoxmlNs.Wp + "effectExtent") ??
            anchor.Element(OoxmlNs.Wp + "extent");
        if (insertAfter is null)
        {
            anchor.AddFirst(wrap);
        }
        else
        {
            insertAfter.AddAfterSelf(wrap);
        }
    }

    private static void SetImageWrapDistance(XElement anchor, string attributeName, long? value)
    {
        if (value is not null)
        {
            anchor.SetAttributeValue(attributeName, value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }

    private static bool TryReadImagePositionAxis(
        DocxPatchOperation operation,
        string axis,
        List<DocxDiagnostic> diagnostics,
        out ImagePositionAxis position)
    {
        position = default;
        bool hasRelative = operation.Fields.TryGetValue($"{axis}-relative", out string? relative);
        bool hasOffset = operation.Fields.TryGetValue($"{axis}-offset", out string? offsetText);
        bool hasAlign = operation.Fields.TryGetValue($"{axis}-align", out string? align);
        if (!hasRelative && !hasOffset && !hasAlign)
        {
            return true;
        }

        if (hasOffset && hasAlign)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5210", $"Image position axis '{axis}' cannot specify both offset and align.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        if (relative is not null && !IsValidImagePositionRelative(axis, relative))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5210", $"Unsupported image {axis} relative value '{relative}'.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        long? offsetEmus = null;
        if (offsetText is not null)
        {
            if (!TryParseSignedDimension(offsetText, out long parsedOffset))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5210", $"Image position field '{axis}-offset' must be a signed dimension.", operation, operation.Fields.GetValueOrDefault("target")));
                return false;
            }

            offsetEmus = parsedOffset;
        }

        if (align is not null && !IsValidImagePositionAlign(axis, align))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5210", $"Unsupported image {axis} align value '{align}'.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        position = new ImagePositionAxis(true, hasRelative ? relative : null, offsetEmus, hasAlign ? align : null);
        return true;
    }

    private static bool IsValidImagePositionRelative(string axis, string value)
    {
        return axis == "horizontal"
            ? value is "page" or "margin" or "column" or "character" or "leftMargin" or "rightMargin" or "insideMargin" or "outsideMargin"
            : value is "page" or "margin" or "paragraph" or "line" or "topMargin" or "bottomMargin" or "insideMargin" or "outsideMargin";
    }

    private static bool IsValidImagePositionAlign(string axis, string value)
    {
        return axis == "horizontal"
            ? value is "left" or "center" or "right" or "inside" or "outside"
            : value is "top" or "center" or "bottom" or "inside" or "outside";
    }

    private static bool TryParseSignedDimension(string text, out long emus)
    {
        emus = 0;
        string trimmed = text.Trim();
        string[] suffixes = ["emu", "in", "cm", "pt", "px"];
        foreach (string suffix in suffixes)
        {
            if (!trimmed.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string numberText = trimmed[..^suffix.Length];
            if (!double.TryParse(numberText, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value))
            {
                return false;
            }

            emus = suffix.ToLowerInvariant() switch
            {
                "emu" => checked((long)Math.Round(value, MidpointRounding.AwayFromZero)),
                "in" => OoxmlUnits.InchesToEmu(value),
                "cm" => OoxmlUnits.CentimetersToEmu(value),
                "pt" => OoxmlUnits.PointsToEmu(value),
                "px" => OoxmlUnits.PixelsToEmu(value),
                _ => 0
            };
            return true;
        }

        return false;
    }

    private static void SetImagePositionAxis(XElement anchor, string elementName, ImagePositionAxis axis)
    {
        if (!axis.HasAny)
        {
            return;
        }

        XElement position = anchor.Element(OoxmlNs.Wp + elementName) ?? new XElement(OoxmlNs.Wp + elementName);
        if (position.Parent is null)
        {
            AddImagePositionElement(anchor, elementName, position);
        }

        if (axis.RelativeFrom is not null)
        {
            position.SetAttributeValue("relativeFrom", axis.RelativeFrom);
        }
        else if (position.Attribute("relativeFrom") is null)
        {
            position.SetAttributeValue("relativeFrom", elementName == "positionH" ? "column" : "paragraph");
        }

        if (axis.OffsetEmus is not null)
        {
            position.Elements(OoxmlNs.Wp + "align").Remove();
            XElement offset = position.Element(OoxmlNs.Wp + "posOffset") ?? new XElement(OoxmlNs.Wp + "posOffset");
            if (offset.Parent is null)
            {
                position.Add(offset);
            }

            offset.Value = axis.OffsetEmus.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        if (axis.Align is not null)
        {
            position.Elements(OoxmlNs.Wp + "posOffset").Remove();
            XElement align = position.Element(OoxmlNs.Wp + "align") ?? new XElement(OoxmlNs.Wp + "align");
            if (align.Parent is null)
            {
                position.Add(align);
            }

            align.Value = axis.Align;
        }

        if (!position.Elements(OoxmlNs.Wp + "posOffset").Any() &&
            !position.Elements(OoxmlNs.Wp + "align").Any())
        {
            position.Add(new XElement(OoxmlNs.Wp + "posOffset", "0"));
        }
    }

    private static void AddImagePositionElement(XElement anchor, string elementName, XElement position)
    {
        if (elementName == "positionH")
        {
            XElement? before = anchor.Element(OoxmlNs.Wp + "positionV") ?? anchor.Element(OoxmlNs.Wp + "extent");
            if (before is null)
            {
                anchor.AddFirst(position);
            }
            else
            {
                before.AddBeforeSelf(position);
            }

            return;
        }

        XElement? horizontal = anchor.Element(OoxmlNs.Wp + "positionH");
        if (horizontal is not null)
        {
            horizontal.AddAfterSelf(position);
            return;
        }

        XElement? extent = anchor.Element(OoxmlNs.Wp + "extent");
        if (extent is null)
        {
            anchor.AddFirst(position);
        }
        else
        {
            extent.AddBeforeSelf(position);
        }
    }

    private static ImageCrop ReadImageCrop(XElement? sourceRectangle)
    {
        return new ImageCrop(
            ReadCropPerThousandPercent(sourceRectangle, "l"),
            ReadCropPerThousandPercent(sourceRectangle, "t"),
            ReadCropPerThousandPercent(sourceRectangle, "r"),
            ReadCropPerThousandPercent(sourceRectangle, "b"));
    }

    private static int ReadCropPerThousandPercent(XElement? sourceRectangle, string localName)
    {
        string? value = (string?)sourceRectangle?.Attribute(localName);
        if (string.IsNullOrWhiteSpace(value))
        {
            return 0;
        }

        if (value.EndsWith("%", StringComparison.Ordinal) &&
            decimal.TryParse(value[..^1], System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out decimal percent))
        {
            return PercentToPerThousand(percent);
        }

        return int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int perThousandPercent)
            ? Math.Clamp(perThousandPercent, 0, 100_000)
            : 0;
    }

    private static bool TryReadCropPercentField(
        DocxPatchOperation operation,
        string fieldName,
        int currentValue,
        List<DocxDiagnostic> diagnostics,
        out int value)
    {
        value = currentValue;
        if (!operation.Fields.TryGetValue(fieldName, out string? text))
        {
            return true;
        }

        if (!decimal.TryParse(text, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out decimal percent) ||
            percent is < 0m or > 100m)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5208", $"Image crop field '{fieldName}' must be a percentage from 0 to 100.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        value = PercentToPerThousand(percent);
        return true;
    }

    private static int PercentToPerThousand(decimal percent)
    {
        return (int)Math.Round(percent * 1000m, MidpointRounding.AwayFromZero);
    }

    private static void SetImageCrop(XElement blipFill, ImageCrop crop)
    {
        XElement? sourceRectangle = blipFill.Element(OoxmlNs.A + "srcRect");
        if (crop.Left == 0 && crop.Top == 0 && crop.Right == 0 && crop.Bottom == 0)
        {
            sourceRectangle?.Remove();
            return;
        }

        if (sourceRectangle is null)
        {
            sourceRectangle = new XElement(OoxmlNs.A + "srcRect");
            XElement? blip = blipFill.Element(OoxmlNs.A + "blip");
            if (blip is null)
            {
                blipFill.AddFirst(sourceRectangle);
            }
            else
            {
                blip.AddAfterSelf(sourceRectangle);
            }
        }

        SetCropAttribute(sourceRectangle, "l", crop.Left);
        SetCropAttribute(sourceRectangle, "t", crop.Top);
        SetCropAttribute(sourceRectangle, "r", crop.Right);
        SetCropAttribute(sourceRectangle, "b", crop.Bottom);
    }

    private static void SetCropAttribute(XElement sourceRectangle, string localName, int value)
    {
        sourceRectangle.SetAttributeValue(localName, value == 0 ? null : value.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteImage(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported delete-image target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        XElement? drawing = imageTarget.Blip.Ancestors(OoxmlNs.W + "drawing").FirstOrDefault();
        if (drawing is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' does not have an editable DrawingML object.", operation, target)];
        }


        drawing.Remove();
        if (!UsesRelationship(imageTarget.Document, imageTarget.RelationshipId))
        {
            package.RemoveRelationship(imageTarget.PartName, imageTarget.RelationshipId, cancellationToken);
            if (!AnyRelationshipTargetsPart(package, imageTarget.Part.Name, cancellationToken))
            {
                package.RemovePart(imageTarget.Part.Name, cancellationToken);
            }
        }

        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static bool UsesRelationship(XDocument document, string relationshipId)
    {
        return document
            .Descendants(OoxmlNs.A + "blip")
            .Any(blip => string.Equals((string?)blip.Attribute(OoxmlNs.R + "embed"), relationshipId, StringComparison.Ordinal));
    }

    private static bool AnyRelationshipTargetsPart(OoxmlPackage package, string partName, CancellationToken cancellationToken)
    {
        foreach (OoxmlPart relationshipPart in package.Parts.Values.Where(part => part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            string sourcePartName = OoxmlPath.GetSourcePartNameFromRelationshipPartName(relationshipPart.Name);
            foreach (OoxmlRelationship relationship in package.GetRelationships(sourcePartName, cancellationToken))
            {
                if (!relationship.IsExternal &&
                    relationship.ResolvedTarget is not null &&
                    string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        return false;
    }
    private static bool ValidateImageContentTypeGuard(
        DocxPatchOperation operation,
        string target,
        string? actualContentType,
        List<DocxDiagnostic> diagnostics)
    {
        if (!operation.Fields.TryGetValue("expect-content-type", out string? expectedContentType))
        {
            return true;
        }

        string? normalizedExpected = NormalizeImageContentType(expectedContentType);
        if (normalizedExpected is null)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'expect-content-type' must be image/png or image/jpeg.", operation, target));
            return false;
        }

        if (!string.Equals(actualContentType, normalizedExpected, StringComparison.Ordinal))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected image content type '{normalizedExpected}', found '{actualContentType ?? "unknown"}'.", operation, target, fieldName: "expect-content-type"));
        }

        return diagnostics.Count == 0;
    }
}
