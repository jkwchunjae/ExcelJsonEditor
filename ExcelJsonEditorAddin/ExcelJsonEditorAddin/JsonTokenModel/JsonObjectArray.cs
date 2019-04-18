using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using Utils;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public class JsonObjectArray : IJsonToken
    {
        private JArray _token;
        private Excel.Worksheet _sheet = null;
        private List<CellData> _cellDatas = new List<CellData>();

        private readonly int _titleRow = 1;

        public JsonTokenType Type() => JsonTokenType.ObjectArray;
        public JToken GetToken() => _token;
        public string Path() => GetToken()?.Path;

        public JsonObjectArray(JArray jArray)
        {
            _token = jArray;
        }

        public void Spread(Excel.Worksheet sheet)
        {
            _sheet = sheet;

            _cellDatas = MakeTitle(_sheet, _token).ToList();
            FillCellData(_sheet, _token, _cellDatas);


            SetNamedRange(sheet, _cellDatas.Where(x => x.Type == DataType.Title));

            _cellDatas.ForEach(x => x.Value?.Spread(x.Cell));
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
            if (cellData.Value.Type() == JsonTokenType.Title)
            {
                Globals.ThisAddIn.Application.EnableEvents = false;
                if (cellData.Key.Extended)
                {
                    var left = cellData.Cell.Offset[0, 1];
                    var right = cellData.Cell.Offset[0, cellData.Key.ColumnCount - 1];

                    var childs = cellData.Key.ChildTitles;
                    _cellDatas = _cellDatas.Where(x => !childs.Contains(x.Key)).ToList();
                    childs.ForEach(x => cellData.Key.RemoveChildTitle(x));

                    _sheet.Range[left, right].EntireColumn.Delete();
                }
                else
                {
                    var titles = _cellDatas.Where(x => x.Cell.Column == cellData.Cell.Column)
                        .Where(x => x.Type == DataType.Value)
                        .Select(x => (JsonObject)x.Value)
                        .SelectMany(x => x.Keys)
                        .Distinct()
                        .Select((x, i) => new
                        {
                            Column = cellData.Cell.Column + i + 1,
                            Path = $"{cellData.Key.Title}.{x}",
                        })
                        .ToList();

                    var left = cellData.Cell.Offset[0, 1];
                    var right = cellData.Cell.Offset[0, titles.Count()];
                    _sheet.Range[left, right].EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                    var titleCellDatas = titles
                        .Select(x => new
                        {
                            Cell = _sheet.Cells[1, x.Column],
                            JsonTitle = new JsonTitle(x.Path),
                        })
                        .Select(x => new CellData
                        {
                            Type = DataType.Title,
                            Cell = x.Cell,
                            Key = x.JsonTitle,
                            Value = x.JsonTitle,
                        })
                        .ToList();

                    titleCellDatas.ForEach(x => x.Key.Spread(x.Cell));
                    titleCellDatas.ForEach(x => cellData.Key.AddChildTitle(x.Key));
                    _cellDatas.AddRange(titleCellDatas);
                    FillCellData(_sheet, _token, _cellDatas);
                }
                Globals.ThisAddIn.Application.EnableEvents = true;
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

        private List<CellData> MakeTitle(Excel.Worksheet sheet, JArray token)
        {
            var titles = token.Where(x => x.Type == JTokenType.Object)
                .Select(x => (JsonObject)x.CreateJsonToken())
                .SelectMany(x => x.Keys)
                .GroupBy(x => x)
                .Select(x => new JsonTitle(x.First()))
                .Select((x, i) => new
                {
                    Index = i,
                    JsonTitle = x,
                    Cell = sheet.Cells[_titleRow, i + 1],
                })
                .Select(x => new CellData
                {
                    Type = DataType.Title,
                    Cell = x.Cell,
                    Index = x.Index,
                    Key = x.JsonTitle,
                    Value = x.JsonTitle,
                })
                .ToList();

            return titles;
        }

        private void FillCellData(Excel.Worksheet sheet, JArray token, List<CellData> cellDatas)
        {
            var titles = cellDatas.Where(x => x.Type == DataType.Title).ToList();

            var objects = token.Where(x => x.Type == JTokenType.Object)
                .Select((x, i) => new { Row = _titleRow + i + 1, JsonObject = (JsonObject)x.CreateJsonToken() })
                .ToList();

            var newDatas = titles.Join(objects, a => 1, b => 1, (title, obj) => new { Title = title, Object = obj })
                .Select(x => new
                {
                    x.Object.Row,
                    x.Title.Cell.Column,
                    x.Title,
                    JsonTitle = (JsonTitle)x.Title.Key,
                    Object = x.Object.JsonObject,
                })
                .Where(x => cellDatas.Empty(e => e.Cell.Row == x.Row && e.Cell.Column == x.Column))
                .Select(x => new CellData
                {
                    Type = DataType.Value,
                    Cell = sheet.Cells[x.Row, x.Column],
                    Key = x.JsonTitle,
                    Value = x.Object.GetToken().SelectToken(x.JsonTitle.Title)?.CreateJsonToken(),
                })
                .ToList();

            newDatas.ForEach(x => x.Value?.Spread(x.Cell));
            cellDatas.AddRange(newDatas);
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
