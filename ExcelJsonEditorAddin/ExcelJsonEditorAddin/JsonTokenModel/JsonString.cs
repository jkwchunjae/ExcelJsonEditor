using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public class JsonString : IJsonToken
    {
        private JValue _token;

        public JsonTokenType Type() => _token.Type.ConvertToJsonTokenType();
        public JToken GetToken() => _token;

        public JsonString(JValue jValue)
        {
            _token = jValue;

        }

        public void Dump(Excel.Worksheet ws)
        {
        }

        public void Dump(Excel.Range cell)
        {
            cell.Value2 = (string)_token.Value;
            if (cell.Value2.ToString() != (string) _token.Value)
            {
                cell.Value2 = "'" + (string) _token.Value;
            }
        }

        public bool OnDoubleClick(Excel.Range target)
        {
            return true;
        }

        public bool OnRightClick(Excel.Range target)
        {
            return false;
        }

        public void OnChangeValue(Excel.Range target)
        {
            _token.Value = target.Value2.ToString();
        }
    }
}
