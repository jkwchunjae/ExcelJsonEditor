using ExcelJsonEditorAddin.Config;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelJsonEditorAddin.Theme;
using System.Drawing;
using ExcelJsonEditorAddin.JsonTokenModel;
using Microsoft.Office.Core;

namespace ExcelJsonEditorAddin
{
    public static class WorkbookExtensions
    {
        public static IEnumerable<Excel.Worksheet> ToEnumerable(this Excel.Sheets sheets)
        {
            for (var i = 1; i <= sheets.Count; i++)
            {
                yield return sheets[i];
            }
        }

        public static IEnumerable<Excel.Style> ToEnumerable(this Excel.Styles styles)
        {
            foreach (Excel.Style style in styles)
            {
                yield return style;
            }
        }

        public static void SaveForJsonEditor(this Excel.Workbook book, string jsonFileName)
        {
            var workbookPath = PathOf.TemporaryFilePath(jsonFileName);

            if (!Directory.Exists(PathOf.LocalRootDirectory))
            {
                Directory.CreateDirectory(PathOf.LocalRootDirectory);
            }

            if (File.Exists(workbookPath))
            {
                try
                {
                    File.Delete(workbookPath);
                }
                catch
                {
                    MessageBox.Show("Opened another file.");
                    throw;
                }
            }

            book.SaveAs(workbookPath);
        }

        public static void ChangeTheme(this Excel.Workbook book, ThemeType themeType)
        {
            var funcDic = new Dictionary<ThemeType, Action<Excel.Workbook>>()
            {
                [ThemeType.White] = ChangeThemeWhite,
                [ThemeType.Dark] = ChangeThemeDark,
            };

            if (!funcDic.ContainsKey(themeType))
            {
                return;
            }

            funcDic[themeType](book);
        }

        private static void ChangeThemeWhite(this Excel.Workbook book)
        {
            var style = book.Styles["Normal"];
            style.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
            style.Font.Color = Color.Black;
            style.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }

        private static void ChangeThemeDark(this Excel.Workbook book)
        {
            var style = book.Styles["Normal"];
            style.Interior.Color = Color.FromArgb(30, 30, 30);
            style.Font.Color = Color.FromArgb(220, 220, 220);

            var borderIndexList = new List<Excel.XlBordersIndex>
            {
                Excel.XlBordersIndex.xlDiagonalDown,
                Excel.XlBordersIndex.xlDiagonalUp,
            };

            style.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            style.Borders.Color = Color.FromArgb(80, 80, 80);
            style.Borders.Weight = Excel.XlBorderWeight.xlThin;

            borderIndexList.ForEach(index =>
            {
                style.Borders[index].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            });
        }

        public static Excel.Workbook SpreadJsonToken(this Excel.Workbook book, Excel.Worksheet currentSheet, IJsonToken jsonToken)
        {
            var sheetName = jsonToken.Path().Replace("[", "").Replace("]", "");

            if (book.Sheets.ToEnumerable().Any(x => x.Name == sheetName))
            {
                Excel.Worksheet sht = book.Sheets[sheetName];
                sht.Activate();
            }
            else
            {
                Excel.Worksheet sheet = book.Sheets.Add(After: currentSheet);
                sheet.Name = sheetName;
                jsonToken.Dump(sheet);
            }
            return book;
        }
    }
}
