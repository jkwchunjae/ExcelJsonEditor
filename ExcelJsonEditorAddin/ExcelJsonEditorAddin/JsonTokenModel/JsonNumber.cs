using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public class JsonNumber : IJsonToken
    {
        private JValue _token;

        public JsonTokenType Type() => _token.Type.ConvertToJsonTokenType();
        public JToken GetToken() => _token;
        public string Path() => GetToken()?.Path;

        public JsonNumber(JValue jValue)
        {
            _token = jValue;
        }

        public void Dump(Excel.Worksheet ws)
        {
        }

        public void Dump(Excel.Range cell)
        {
            cell.Value2 = _token.Value;
        }

        public bool OnDoubleClick(Excel.Workbook book, Excel.Range target)
        {
            return true;
        }

        public bool OnRightClick(Excel.Range target)
        {
            return false;
        }

        public void OnChangeValue(Excel.Range target)
        {
            string targetString = target.Value2.ToString();
            if (int.TryParse(targetString, out var intNumber))
            {
                _token.Value = intNumber;
            }
            else if (double.TryParse(targetString, out var floatNumber))
            {
                _token.Value = floatNumber;
            }
        }
    }
}
