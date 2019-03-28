using ExcelJsonEditorAddin.JsonTokenModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.EditorData
{
    public class BookData
    {
        public string WorkbookName { get; set; }
        public Excel.Workbook Workbook { get; set; }
        public IJsonToken RootJsonToken { get; set; }
        public string JsonPath { get; set; }
    }
}
