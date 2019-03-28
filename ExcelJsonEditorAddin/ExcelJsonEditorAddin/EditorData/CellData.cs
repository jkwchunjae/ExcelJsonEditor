using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public enum DataType
    {
        Key,
        Value,
        Title,
    }

    public class CellData
    {
        public DataType Type { get; set; }
        public string Address { get; set; }
        public Excel.Range Cell { get; set; }
        public int Index { get; set; }
        public IJsonToken Key { get; set; }
        public IJsonToken Value { get; set; }
    }
}
