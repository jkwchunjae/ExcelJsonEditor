using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public class JsonValue : IJsonToken
    {
        private JValue _token;

        public JsonTokenType Type() => JsonTokenType.Other;
        public JToken GetToken() => _token;
        public string Path() => GetToken()?.Path;

        public JsonValue(JValue jValue)
        {
            _token = jValue;
        }

        public void Spread(Excel.Worksheet ws)
        {
        }

        public void Spread(Excel.Range cell)
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
            _token.Value = target.Value2;
        }
    }

}
