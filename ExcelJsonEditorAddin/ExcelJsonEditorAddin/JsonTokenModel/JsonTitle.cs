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
        private List<JsonTitle> _childTitles = new List<JsonTitle>();
        public string Title { get; private set; }

        public JsonTokenType Type() => JsonTokenType.Title;
        public JToken GetToken() => null;
        public string Path() => GetToken()?.Path;

        public bool Extended => _childTitles.Any();
        public int ColumnCount => 1 + _childTitles.Sum(x => x.ColumnCount);
        public List<JsonTitle> ChildTitles => _childTitles
            .Concat(_childTitles.SelectMany(e => e.ChildTitles))
            .ToList();

        public JsonTitle(string path)
        {
            Title = path;
        }

        public void AddChildTitle(JsonTitle title)
        {
            _childTitles.Add(title);
        }

        public void RemoveChildTitle(JsonTitle title)
        {
            if (_childTitles.Contains(title))
            {
                _childTitles.Remove(title);
            }
            else
            {
                _childTitles.ForEach(x => x.RemoveChildTitle(title));
            }
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
