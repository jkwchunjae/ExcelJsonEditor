using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public class JsonTitle : IJsonToken
    {
        private List<JProperty> _tokens;
        public string Title => _tokens.FirstOrDefault()?.Name;

        public JsonTokenType Type() => JsonTokenType.Title;
        public JToken GetToken() => null;
        public string Path() => GetToken()?.Path;

        public JsonTitle(IEnumerable<JProperty> jProperties)
        {
            _tokens = jProperties.ToList();
        }

        public void Spread(Excel.Worksheet ws)
        {
        }

        public void Spread(Excel.Range cell)
        {
            cell.Value2 = Title;
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
        }
    }
}
