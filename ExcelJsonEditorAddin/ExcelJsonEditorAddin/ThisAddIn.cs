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
using System.Windows.Forms;
using System.Drawing;
using ExcelJsonEditorAddin.Theme;
using ExcelJsonEditorAddin.Config;
using ExcelJsonEditorAddin.Extensions;

namespace ExcelJsonEditorAddin
{
    public partial class ThisAddIn
    {
        private StartupControl _startupControl;
        private CustomTaskPane _customTaskPane;

        private List<Excel.Workbook> _workbookList = new List<Excel.Workbook>();
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
                _workbookList.ForEach(x => x.ChangeTheme(settings.Theme));
            }
            _settings = settings;
            _settings.Save();
        }

        private void _startupControl_OpenFiles(object sender, string filePath)
        {
            Excel.Workbook book = Application.Workbooks.Add();

            try {
                book.Initialize(filePath, _settings);
                _workbookList.Add(book);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                book.Close(SaveChanges: false);
            }
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
