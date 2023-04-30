using System;
using System.Text.Json.Serialization;
using System.Text.Json;

namespace Event_Server.Network
{
    public class TupleConverter : JsonConverter<(byte, string, int)>
    {
        public override void Write(Utf8JsonWriter writer, (byte, string, int) value, JsonSerializerOptions options)
        {
            writer.WriteStartArray();
            writer.WriteNumberValue(value.Item1);
            writer.WriteStringValue(value.Item2);
            writer.WriteNumberValue(value.Item3);
            writer.WriteEndArray();
        }

        public override (byte, string, int) Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            reader.Read(); // StartArray
            var item1 = reader.GetByte();
            reader.Read(); // String
            var item2 = reader.GetString();
            reader.Read(); // Number
            var item3 = reader.GetInt32();
            reader.Read(); // EndArray
            return (item1, item2, item3);
        }
    }
}
