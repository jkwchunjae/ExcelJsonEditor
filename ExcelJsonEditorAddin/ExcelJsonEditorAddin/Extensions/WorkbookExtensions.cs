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
using ExcelJsonEditorAddin.Extensions;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using ExcelJsonEditorAddin.EditorData;
using Utils;

namespace ExcelJsonEditorAddin
{
    public static class WorkbookExtensions
    {
        private static List<BookData> _bookDatas = new List<BookData>();

        public static IEnumerable<Excel.Worksheet> SheetList(this Excel.Workbook book)
        {
            foreach (Excel.Worksheet sheet in book.Sheets)
            {
                yield return sheet;
            }
        }

        public static IEnumerable<Excel.Style> StyleList(this Excel.Workbook book)
        {
            foreach (Excel.Style style in book.Styles)
            {
                yield return style;
            }
        }

        public static Excel.Workbook Initialize(this Excel.Workbook book, string jsonFilePath, Settings settings)
        {
            var fileName = Path.GetFileNameWithoutExtension(jsonFilePath);
            var jtoken = JsonConvert.DeserializeObject<JToken>(File.ReadAllText(jsonFilePath, Encoding.UTF8));
            var jsonToken = jtoken.CreateJsonToken();

            book.ChangeTheme(settings.Theme);

            Excel.Worksheet sheet = book.SheetList().First();
            sheet.Spread(jsonToken, fileName);

            book.SaveForJsonEditor(fileName);

            _bookDatas.Add(new BookData
            {
                WorkbookName = $"{fileName}.xlsx",
                RootJsonToken = jsonToken,
                Workbook = book,
                JsonPath = jsonFilePath,
            });

            book.SheetBeforeDoubleClick += Book_SheetBeforeDoubleClick;
            book.SheetBeforeRightClick += Book_SheetBeforeRightClick;
            book.AfterSave += Book_AfterSave;

            return book;
        }

        private static void Book_SheetBeforeDoubleClick(object sh, Excel.Range target, ref bool cancel)
        {
            var book = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (_bookDatas.Any(x => x.WorkbookName == book.Name))
            {
                var bookData = _bookDatas.First(x => x.WorkbookName == book.Name);
                cancel = bookData.RootJsonToken.OnDoubleClick(bookData.Workbook, target);
            }
        }

        private static void Book_SheetBeforeRightClick(object sh, Excel.Range target, ref bool cancel)
        {
            var book = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (_bookDatas.Any(x => x.WorkbookName == book.Name))
            {
                var bookData = _bookDatas.First(x => x.WorkbookName == book.Name);
                cancel = bookData.RootJsonToken.OnRightClick(target);
            }
        }

        private static void Book_AfterSave(bool success)
        {
            if (!success)
            {
                return;
            }

            var book = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (_bookDatas.Empty(x => x.WorkbookName == book.Name))
            {
                MessageBox.Show($"Unknown workbook. (WorkbookName: {book.Name})");
                return;
            }
            var bookData = _bookDatas.First(x => x.WorkbookName == book.Name);
            File.WriteAllText(bookData.JsonPath, bookData.RootJsonToken.GetToken().Serialize(Formatting.Indented), Encoding.UTF8);
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
            var normalStyle = book.Styles[StyleName.Normal];
            normalStyle.SetDefaultStyle(themeType);

            book.CreateStyle(StyleName.Number).SetNumberStyle(themeType);
            book.CreateStyle(StyleName.String).SetStringStyle(themeType);
        }

        public static Excel.Style CreateStyle(this Excel.Workbook book, string styleName)
        {
            if (book.StyleList().Empty(x => x.Name == styleName))
            {
                book.Styles.Add(styleName);
            }
            return book.Styles[styleName];
        }

        public static Excel.Workbook SpreadJsonToken(this Excel.Workbook book, Excel.Worksheet currentSheet, IJsonToken jsonToken)
        {
            var sheetName = jsonToken.Path().ConvertSheetName();

            if (book.SheetList().Any(x => x.Name == sheetName))
            {
                Excel.Worksheet sht = book.Sheets[sheetName];
                sht.Activate();
            }
            else
            {
                Excel.Worksheet sheet = book.Sheets.Add(After: currentSheet);
                sheet.Spread(jsonToken);
            }
            return book;
        }
    }
}
