using ExcelJsonEditorAddin.Theme;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJsonEditorAddin.Config
{
    public class Settings
    {
        public ThemeType Theme { get; set; } = ThemeType.White;

        public static Settings Open()
        {
            if (!File.Exists(PathOf.SettingsPath))
            {
                return new Settings();
            }
            var jsonText = File.ReadAllText(PathOf.SettingsPath, Encoding.UTF8);
            return JsonConvert.DeserializeObject<Settings>(jsonText);
        }

        public void Save()
        {
            var jsonText = JsonConvert.SerializeObject(this, Formatting.Indented);
            if (!Directory.Exists(PathOf.LocalRootDirectory))
            {
                Directory.CreateDirectory(PathOf.LocalRootDirectory);
            }
            File.WriteAllText(PathOf.SettingsPath, jsonText, Encoding.UTF8);
        }
    }
}
