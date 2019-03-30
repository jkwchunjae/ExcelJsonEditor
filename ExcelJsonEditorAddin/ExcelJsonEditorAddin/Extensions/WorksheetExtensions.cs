using ExcelJsonEditorAddin.JsonTokenModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.Extensions
{
    public static class WorksheetExtensions
    {
        public static Excel.Worksheet Spread(this Excel.Worksheet sheet, IJsonToken jsonToken, string sheetName = null)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            sheet.Name = (sheetName ?? jsonToken.Path()).ConvertSheetName();
            jsonToken.Spread(sheet);
            sheet.Change += jsonToken.OnChangeValue;
            //sheet.Protect();

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            return sheet;
        }

        public static string ConvertSheetName(this string str)
        {
            return str
                .Replace("[", "")
                .Replace("]", "");
        }
    }
}
