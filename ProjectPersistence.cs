using System;
using System.Text.Json;
using System.Text.Json.Serialization;
using Windows.Foundation;

namespace SnapSlate;

internal static class ProjectPersistence
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        Converters =
        {
            new JsonStringEnumConverter(JsonNamingPolicy.CamelCase),
            new PointJsonConverter(),
            new RectJsonConverter()
        }
    };

    public static string Serialize(SnapSlateProjectState state)
    {
        return JsonSerializer.Serialize(state, SerializerOptions);
    }

    public static SnapSlateProjectState? Deserialize(string json)
    {
        return JsonSerializer.Deserialize<SnapSlateProjectState>(json, SerializerOptions);
    }

    private sealed class PointJsonConverter : JsonConverter<Point>
    {
        public override Point Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (reader.TokenType != JsonTokenType.StartObject)
            {
                throw new JsonException("Point must be written as an object.");
            }

            var x = 0d;
            var y = 0d;

            while (reader.Read())
            {
                if (reader.TokenType == JsonTokenType.EndObject)
                {
                    return new Point(x, y);
                }

                if (reader.TokenType != JsonTokenType.PropertyName)
                {
                    throw new JsonException("Unexpected token while reading Point.");
                }

                var propertyName = reader.GetString();
                if (!reader.Read())
                {
                    throw new JsonException("Unexpected end of Point data.");
                }

                switch (propertyName)
                {
                    case "x":
                        x = reader.GetDouble();
                        break;
                    case "y":
                        y = reader.GetDouble();
                        break;
                }
            }

            throw new JsonException("Unexpected end of Point data.");
        }

        public override void Write(Utf8JsonWriter writer, Point value, JsonSerializerOptions options)
        {
            writer.WriteStartObject();
            writer.WriteNumber("x", value.X);
            writer.WriteNumber("y", value.Y);
            writer.WriteEndObject();
        }
    }

    private sealed class RectJsonConverter : JsonConverter<Rect>
    {
        public override Rect Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (reader.TokenType != JsonTokenType.StartObject)
            {
                throw new JsonException("Rect must be written as an object.");
            }

            var x = 0d;
            var y = 0d;
            var width = 0d;
            var height = 0d;

            while (reader.Read())
            {
                if (reader.TokenType == JsonTokenType.EndObject)
                {
                    return new Rect(x, y, width, height);
                }

                if (reader.TokenType != JsonTokenType.PropertyName)
                {
                    throw new JsonException("Unexpected token while reading Rect.");
                }

                var propertyName = reader.GetString();
                if (!reader.Read())
                {
                    throw new JsonException("Unexpected end of Rect data.");
                }

                switch (propertyName)
                {
                    case "x":
                        x = reader.GetDouble();
                        break;
                    case "y":
                        y = reader.GetDouble();
                        break;
                    case "width":
                        width = reader.GetDouble();
                        break;
                    case "height":
                        height = reader.GetDouble();
                        break;
                }
            }

            throw new JsonException("Unexpected end of Rect data.");
        }

        public override void Write(Utf8JsonWriter writer, Rect value, JsonSerializerOptions options)
        {
            writer.WriteStartObject();
            writer.WriteNumber("x", value.X);
            writer.WriteNumber("y", value.Y);
            writer.WriteNumber("width", value.Width);
            writer.WriteNumber("height", value.Height);
            writer.WriteEndObject();
        }
    }
}
