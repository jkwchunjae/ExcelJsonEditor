using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;
using ExcelJsonEditorAddin.EditorData;
using Utils;
using System.Windows.Forms;
using System.Drawing;
using ExcelJsonEditorAddin.Theme;
using ExcelJsonEditorAddin.Config;

namespace ExcelJsonEditorAddin
{
    public partial class ThisAddIn
    {
        private StartupControl _startupControl;
        private CustomTaskPane _customTaskPane;

        private List<BookData> _bookDatas = new List<BookData>();

        private Settings _settings = new Settings();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _settings = Settings.Open();

            _startupControl = new StartupControl();
            _startupControl.OpenFiles += _startupControl_OpenFiles;
            _customTaskPane = this.CustomTaskPanes.Add(_startupControl, "Excel Json Editor");
            _customTaskPane.Visible = true;

            if (_settings.SetDefaultExcelTheme)
            {
                Application.ActiveWorkbook.ChangeTheme(_settings.Theme);
            }

            var ribbon = Globals.Ribbons.GetRibbon<ExcelJsonEditorRibbon>();
            ribbon.ChangeSettings += Ribbon_ChangeSettings;
            ribbon.Initialize(_settings);
        }

        private void Ribbon_ChangeSettings(object sender, Settings settings)
        {
            if (_settings.Theme != settings.Theme)
            {
                _bookDatas.ForEach(x => x.Workbook.ChangeTheme(settings.Theme));
            }
            _settings = settings;
            _settings.Save();
        }

        private void _startupControl_OpenFiles(object sender, string filePath)
        {
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var jtoken = JsonConvert.DeserializeObject<JToken>(File.ReadAllText(filePath, Encoding.UTF8));
            var token = jtoken.CreateJsonToken();

            Excel.Workbook book = Application.Workbooks.Add();

            book.ChangeTheme(_settings.Theme);

            _bookDatas.Add(new BookData
            {
                WorkbookName = $"{fileName}.xlsx",
                RootJsonToken = token,
                Workbook = book,
                JsonPath = filePath,
            });

            Excel.Worksheet sheet = book.Sheets.ToEnumerable().First();
            sheet.Name = fileName;

            try
            {
                book.SaveForJsonEditor(fileName);
            }
            catch
            {
                book.Close(SaveChanges: false);
                return;
            }

            Application.ScreenUpdating = false;
            token.Dump(sheet);
            sheet.Change += token.OnChangeValue;
            //sheet.Protect();
            Application.ScreenUpdating = true;

            book.AfterSave += Book_AfterSave;
        }

        private void Book_AfterSave(bool success)
        {
            if (!success)
            {
                return;
            }

            var book = Application.ActiveWorkbook;

            if (_bookDatas.Empty(x => x.WorkbookName == book.Name))
            {
                MessageBox.Show($"Unknown workbook. (WorkbookName: {book.Name})");
                return;
            }
            var bookData = _bookDatas.First(x => x.WorkbookName == book.Name);
            File.WriteAllText(bookData.JsonPath, bookData.RootJsonToken.GetToken().Serialize(Formatting.Indented), Encoding.UTF8);
        }

        private void Ws_BeforeDoubleClick(Excel.Range Target, ref bool Cancel)
        {
            Target.Offset[1, 0].Value2 = Target.Value2;
            Cancel = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
