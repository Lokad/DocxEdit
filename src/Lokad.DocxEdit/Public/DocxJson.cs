using System.Text.Json;
using System.Text.Json.Serialization;

namespace Lokad.DocxEdit;

/// <summary>Creates the JSON wire format used by the command line and hosted commands.</summary>
public static class DocxJson
{
    /// <summary>Creates independent options with the canonical wire-value converters.</summary>
    public static JsonSerializerOptions CreateOptions(bool writeIndented)
    {
        return new JsonSerializerOptions
        {
            WriteIndented = writeIndented,
            Converters = { new DocxOrientationJsonConverter(), new DocxTargetStatusJsonConverter(), new DocxTargetSourceJsonConverter(), new DocxTargetReasonJsonConverter(), new DocxRefreshPolicyJsonConverter(), new DocxLabelStatusJsonConverter(), new DocxLabelSourceJsonConverter(), new DocxVerticalMergeJsonConverter(), new TrackChangesModeJsonConverter() }
        };
    }

    private sealed class DocxOrientationJsonConverter : JsonConverter<DocxOrientation>
    {
        public override DocxOrientation Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxOrientationExtensions.TryParseWireValue(value, out DocxOrientation orientation))
            {
                return orientation;
            }

            throw new JsonException($"Unsupported section orientation '{value}'. Expected portrait or landscape.");
        }

        public override void Write(Utf8JsonWriter writer, DocxOrientation value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class TrackChangesModeJsonConverter : JsonConverter<TrackChangesMode>
    {
        public override TrackChangesMode Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (TrackChangesModeExtensions.TryParseWireValue(value, out TrackChangesMode mode))
            {
                return mode;
            }

            throw new JsonException($"Unsupported track-changes mode: {value}. Expected off, preserve, suggest, or require.");
        }

        public override void Write(Utf8JsonWriter writer, TrackChangesMode value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxTargetStatusJsonConverter : JsonConverter<DocxTargetStatus>
    {
        public override DocxTargetStatus Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxTargetStatusExtensions.TryParseWireValue(value, out DocxTargetStatus parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported target status '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxTargetStatus value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxTargetSourceJsonConverter : JsonConverter<DocxTargetSource>
    {
        public override DocxTargetSource Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxTargetSourceExtensions.TryParseWireValue(value, out DocxTargetSource parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported target source '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxTargetSource value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxTargetReasonJsonConverter : JsonConverter<DocxTargetReason>
    {
        public override DocxTargetReason Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxTargetReasonExtensions.TryParseWireValue(value, out DocxTargetReason parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported target reason '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxTargetReason value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxRefreshPolicyJsonConverter : JsonConverter<DocxRefreshPolicy>
    {
        public override DocxRefreshPolicy Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxRefreshPolicyExtensions.TryParseWireValue(value, out DocxRefreshPolicy parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported refresh policy '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxRefreshPolicy value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxLabelStatusJsonConverter : JsonConverter<DocxLabelStatus>
    {
        public override DocxLabelStatus Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxLabelStatusExtensions.TryParseWireValue(value, out DocxLabelStatus parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported label status '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxLabelStatus value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxLabelSourceJsonConverter : JsonConverter<DocxLabelSource>
    {
        public override DocxLabelSource Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxLabelSourceExtensions.TryParseWireValue(value, out DocxLabelSource parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported label source '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxLabelSource value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxVerticalMergeJsonConverter : JsonConverter<DocxVerticalMerge>
    {
        public override DocxVerticalMerge Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxVerticalMergeExtensions.TryParseWireValue(value, out DocxVerticalMerge parsed))
            {
                return parsed;
            }

            throw new JsonException("Unsupported vertical merge value.");
        }

        public override void Write(Utf8JsonWriter writer, DocxVerticalMerge value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

}
