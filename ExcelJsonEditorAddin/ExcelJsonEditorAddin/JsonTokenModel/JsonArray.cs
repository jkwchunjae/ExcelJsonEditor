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

        public JsonTokenType Type() => _token.Type.ConvertToJsonTokenType();
        public JToken GetToken() => _token;
        public string Path() => GetToken()?.Path;

        public JsonArray(JArray jArray)
        {
            _token = jArray;
        }

        public void Dump(Excel.Worksheet sheet)
        {
            _sheet = sheet;
            sheet.Range["A1"].Value2 = _token.Path;

            if (_token.Empty())
            {
                sheet.Range[_titleRow, 1].Value2 = "Value";
                sheet.Range[_titleRow + 1, 1].Value2 = "<<empty>>";
            }
            else
            {
                if (_token[0].Type != JTokenType.Object)
                {
                    sheet.Cells[_titleRow, 1].Value2 = "Value";
                }
                _cellDatas = MakeCellData(_sheet, _token).ToList();

                SetNamedRange(sheet, _cellDatas.Where(x => x.Type == DataType.Title));

                _cellDatas.ForEach(x => x.Value.Dump(x.Cell));
            }
        }

        public void Dump(Excel.Range cell)
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
            var cellData = _cellDatas.FirstOrDefault(x => x.Address == target.Address);

            cellData?.Value.OnChangeValue(target);
        }

        private IEnumerable<CellData> MakeCellData(Excel.Worksheet sheet, JArray token)
        {
            if (token.Empty())
            {
                return new List<CellData>();
            }
            else if (_token[0].Type == JTokenType.Object)
            {
                var titles = _token.Where(x => x.Type == JTokenType.Object)
                    .Select(x => (JsonObject)x.CreateJsonToken())
                    .SelectMany(x => x.Keys)
                    .GroupBy(x => x.Name)
                    .Select(x => new JsonTitle(x))
                    .Select((x, i) => new
                    {
                        Index = i,
                        JsonTitle = x,
                        Cell = (Excel.Range)sheet?.Cells[_titleRow, i + 1],
                    })
                    .Select(x => new CellData
                    {
                        Type = DataType.Title,
                        Address = x.Cell?.Address,
                        Cell = x.Cell,
                        Index = x.Index,
                        Key = x.JsonTitle,
                        Value = x.JsonTitle,
                    })
                    .ToList();

                var columnDic = titles
                    .ToDictionary(x => ((JsonTitle)x.Key).Title, x => new { Column = x.Index + 1, JsonTitle = (JsonTitle)x.Key });

                var datas = _token.Where(x => x.Type == JTokenType.Object)
                    .Select((x, i) => new { Row = _titleRow + i + 1, JsonObject = (JsonObject)x.CreateJsonToken() })
                    .SelectMany(x => x.JsonObject.Keys.Select(e => new { JsonProperty = e, JToken = e.Value, x.Row }))
                    .Select(x => new
                    {
                        x.Row,
                        Column = columnDic[x.JsonProperty.Name].Column,
                        x.JToken,
                        JsonTitle = columnDic[x.JsonProperty.Name].JsonTitle,
                    })
                    .Select(x => new
                    {
                        Cell = (Excel.Range)sheet?.Cells[x.Row, x.Column],
                        x.JToken,
                        x.JsonTitle,
                    })
                    .Select(x => new CellData
                    {
                        Type = DataType.Value,
                        Address = x.Cell?.Address,
                        Cell = x.Cell,
                        Key = x.JsonTitle,
                        Value = x.JToken.CreateJsonToken(),
                    })
                    .ToList();

                return titles.Concat(datas);
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
                    Address = x.Cell?.Address,
                    Cell = x.Cell,
                    Key = null,
                    Value = x.JToken,
                });
            }
        }

        private void SetNamedRange(Excel.Worksheet sheet, IEnumerable<CellData> titles)
        {
            titles.ToList().ForEach(x =>
            {
                var jsonTitle = (JsonTitle)x.Key;
                var columnText = x.Cell.Address.Split('$').FirstOrDefault(e => e.Length > 0);
                sheet.Names.Add(jsonTitle.Title, sheet.Range[$"{columnText}:{columnText}"]);
            });
        }
    }
}
