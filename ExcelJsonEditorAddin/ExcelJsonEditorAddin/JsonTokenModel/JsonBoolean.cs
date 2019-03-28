using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public class JsonBoolean : IJsonToken
    {
        private JValue _token;

        public JsonTokenType Type() => _token.Type.ConvertToJsonTokenType();
        public JToken GetToken() => _token;

        public JsonBoolean(JValue jValue)
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
            if (bool.TryParse((string)target.Value2, out var value))
            {
                _token.Value = value;
            }
            else
            {
                _token.Value = false;
            }
        }
    }

}
