using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJsonEditorAddin.Theme
{
    [JsonConverter(typeof(StringEnumConverter))]
    public enum ThemeType
    {
        White,
        Dark,
    }

    public static class StyleName
    {
        public static string Normal => "Normal";
        public static string Title => "JsonTitle";
        public static string Number => "JsonNumber";
        public static string String => "JsonString";
        public static string Boolean => "JsonBoolean";
        public static string Array => "JsonArray";
        public static string Object => "JsonObject";
    }
}
