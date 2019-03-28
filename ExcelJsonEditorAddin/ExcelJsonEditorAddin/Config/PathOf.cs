using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJsonEditorAddin.Config
{
    public static class PathOf
    {
        private static readonly string ProjectName = "ExcelJsonEditor";

        public static string LocalRootDirectory
            => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ProjectName);

        public static string TemporaryFilePath(string fileName)
            => Path.Combine(LocalRootDirectory, $"{fileName}.xlsx");

        public static string SettingsPath
            => Path.Combine(LocalRootDirectory, "settings.json");
    }
}
