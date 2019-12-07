using ExcelJsonEditorAddin.JsonTokenModel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;

namespace ExcelJsonEditorAddin
{
    public static class JTokenExtensions
    {
        public static IJsonToken CreateJsonToken(this JToken token)
        {
            if (token.Type == JTokenType.Array)
            {
                var arrayToken = (JArray)token;
                if (arrayToken.Any() && arrayToken[0].Type == JTokenType.Object)
                {
                    return new JsonObjectArray(arrayToken);
                }
            }

            switch (token.Type)
            {
                case JTokenType.Array:
                    return new JsonArray((JArray)token);
                case JTokenType.Object:
                    return new JsonObject((JObject)token);
                case JTokenType.Property:
                    return new JsonProperty((JProperty)token);
                case JTokenType.String:
                    return new JsonString((JValue)token);
                case JTokenType.Integer:
                case JTokenType.Float:
                    return new JsonNumber((JValue)token);
                default:
                    return new JsonValue((JValue)token);
            }
        }

        public static string Serialize(this JToken value)
            => JsonConvert.SerializeObject(value);
        public static string Serialize(this JToken value, Newtonsoft.Json.Formatting formatting)
            => JsonConvert.SerializeObject(value, formatting);
        public static string Serialize(this JToken value, params JsonConverter[] converters)
            => JsonConvert.SerializeObject(value, converters);
        public static string Serialize(this JToken value, Newtonsoft.Json.Formatting formatting, params JsonConverter[] converters)
            => JsonConvert.SerializeObject(value, formatting, converters);
        public static string Serialize(this JToken value, JsonSerializerSettings settings)
            => JsonConvert.SerializeObject(value, settings);
        public static string Serialize(this JToken value, Type type, JsonSerializerSettings settings)
            => JsonConvert.SerializeObject(value, type, settings);
        public static string Serialize(this JToken value, Newtonsoft.Json.Formatting formatting, JsonSerializerSettings settings)
            => JsonConvert.SerializeObject(value, formatting, settings);
        public static string Serialize(this JToken value, Type type, Newtonsoft.Json.Formatting formatting, JsonSerializerSettings settings)
            => JsonConvert.SerializeObject(value, type, formatting, settings);
    }
}
