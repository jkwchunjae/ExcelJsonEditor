using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public class JsonProperty : IJsonToken
    {
        private JProperty _token;

        public JsonTokenType Type() => JsonTokenType.Property;
        public JToken GetToken() => _token;
        public string Path() => GetToken()?.Path;

        public JsonProperty(JProperty jValue)
        {
            _token = jValue;
        }

        public object ToValue()
        {
            return _token.Name;
        }

        public void Spread(Excel.Worksheet ws)
        {
        }

        public void Spread(Excel.Range cell)
        {
            cell.Value2 = _token.Name;
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
            string name = target.Value2.ToString();
            ChangeName(name);
        }

        private void ChangeName(string name)
        {
            var value = _token.Value;

            var newProperty = new JProperty(name, value);
            _token.Replace(newProperty);
        }
    }
}
