using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public class JsonObject : IJsonToken
    {
        private JObject _token;
        private Excel.Worksheet _sheet = null;
        private List<CellData> _cellDatas = new List<CellData>();

        private readonly int _titleRow = 1;

        public JsonTokenType Type() => JsonTokenType.Object;
        public JToken GetToken() => _token;
        public string Path() => GetToken()?.Path;

        public IEnumerable<string> Keys => _cellDatas
            .Where(x => x.Type == DataType.Key)
            .Select(x => x.Key.Title);

        public JsonObject(JObject jObject)
        {
            _token = jObject;

            _cellDatas = MakeCellData(null, _token).ToList();
        }

        public void Spread(Excel.Worksheet sheet)
        {
            _sheet = sheet;
            _sheet.Cells[_titleRow, 1].Value2 = "Key";
            _sheet.Cells[_titleRow, 2].Value2 = "Value";

            _cellDatas = MakeCellData(_sheet, _token).ToList();
            SetNamedRange(_sheet, _cellDatas.Where(x => x.Type == DataType.Key));
            _cellDatas.ForEach(x => x.Value.Spread(x.Cell));
        }

        public void Spread(Excel.Range cell)
        {
            cell.Value2 = "{object}";
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

            if (cellData?.Type == DataType.Key)
            {
                _cellDatas = MakeCellData(_sheet, _token).ToList();
            }
        }

        private IEnumerable<CellData> MakeCellData(Excel.Worksheet sheet, JObject token)
            => token.Properties()
                .Select((x, i) => new
                {
                    Index = i,
                    Property = x,
                    PropertyToken = new JsonTitle(x.Name),
                    ValueToken = x.Value.CreateJsonToken(),
                })
                .SelectMany(x => new CellData[]
                {
                    new CellData
                    {
                        Type = DataType.Key,
                        Cell = (Excel.Range)sheet?.Cells[x.Index + _titleRow + 1, 1],
                        Key = x.PropertyToken,
                        Value = x.PropertyToken,
                    },
                    new CellData
                    {
                        Type = DataType.Value,
                        Cell = (Excel.Range)sheet?.Cells[x.Index + _titleRow + 1, 2],
                        Key = x.PropertyToken,
                        Value = x.ValueToken,
                    }
                });

        private void SetNamedRange(Excel.Worksheet sheet, IEnumerable<CellData> keys)
        {
            keys.ToList().ForEach(x =>
            {
                var propertyName = x.Key.Title;
                sheet.Names.Add(propertyName, x.Cell.Offset[0, 1]);
            });
        }
    }
}
