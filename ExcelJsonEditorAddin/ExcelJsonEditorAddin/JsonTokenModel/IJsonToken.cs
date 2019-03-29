using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public interface IJsonToken
    {
        JsonTokenType Type();
        JToken GetToken();
        string Path();

        void Dump(Excel.Worksheet worksheet);
        void Dump(Excel.Range cell);

        bool OnDoubleClick(Excel.Workbook book, Excel.Range target);
        bool OnRightClick(Excel.Range target);
        void OnChangeValue(Excel.Range target);
    }
}
