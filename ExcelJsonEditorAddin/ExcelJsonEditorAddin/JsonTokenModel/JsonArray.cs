using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using Utils;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public class JsonArray : IJsonToken
    {
        private JArray _token;
        private Excel.Worksheet _sheet = null;
        private List<CellData> _cellDatas = new List<CellData>();

        private readonly int _titleRow = 1;

        public JsonTokenType Type() => JsonTokenType.Array;
        public JToken GetToken() => _token;
        public string Path() => GetToken()?.Path;

        public JsonArray(JArray jArray)
        {
            _token = jArray;
        }

        public void Spread(Excel.Worksheet sheet)
        {
            _sheet = sheet;
            sheet.Cells[_titleRow, 1].Value2 = "Value";

            if (_token.Empty())
            {
                sheet.Cells[_titleRow + 1, 1].Value2 = "<<empty>>";
            }
            else
            {
                _cellDatas = MakeCellData(_sheet, _token).ToList();

                _cellDatas.ForEach(x => x.Value.Spread(x.Cell));
            }
        }

        public void Spread(Excel.Range cell)
        {
            cell.Value2 = "[array]";
        }

        public bool OnDoubleClick(Excel.Workbook book, Excel.Range target)
        {
            if (_cellDatas.Empty(x => x.Cell.Address == target.Address))
            {
                return false;
            }

            var cellData = _cellDatas.First(x => x.Cell.Address == target.Address);
            if (cellData.Value.CanSpreadType())
            {
                book.SpreadJsonToken(_sheet, cellData.Value);
            }
            return true;
        }

        public bool OnRightClick(Excel.Range target)
        {
            return false;
        }

        public void OnChangeValue(Excel.Range target)
        {
            var cellData = _cellDatas.FirstOrDefault(x => x.Cell.Address == target.Address);

            cellData?.Value.OnChangeValue(target);
        }

        private IEnumerable<CellData> MakeCellData(Excel.Worksheet sheet, JArray token)
        {
            if (token.Empty())
            {
                return new List<CellData>();
            }
            else
            {
                return _token.Select((x, i) => new
                {
                    Index = i,
                    Cell = (Excel.Range)sheet?.Cells[_titleRow + i + 1, 1],
                    JToken = x.CreateJsonToken(),
                })
                .Select(x => new CellData
                {
                    Type = DataType.Value,
                    Index = x.Index,
                    Cell = x.Cell,
                    Key = null,
                    Value = x.JToken,
                });
            }
        }
    }
}
