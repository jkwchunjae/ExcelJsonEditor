using System;
using System.Collections.Generic;
using System.Linq;
using JkwExtensions;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.JsonTokenModel
{
    public class JsonObjectArray : IJsonToken
    {
        private JArray _token;
        private Excel.Worksheet _sheet = null;
        private List<CellData> _cellDataList = new List<CellData>();

        private readonly int _titleRow = 1;

        public JsonTokenType Type() => JsonTokenType.ObjectArray;
        public JToken GetToken() => _token;
        public string Path() => GetToken()?.Path;

        public JsonObjectArray(JArray jArray)
        {
            _token = jArray;
        }

        public object ToValue()
        {
            return "[array]";
        }

        public void Spread(Excel.Worksheet sheet)
        {
            _sheet = sheet;

            var tokenArray = _token;

            var objectList = tokenArray
                .Where(x => x.Type == JTokenType.Object)
                .Select((x, i) => new { Index = i, Object = (JsonObject)x.CreateJsonToken() })
                .ToList();

            var titleColumnDic = objectList
                .SelectMany(x => x.Object.Keys)
                .Distinct()
                .Select((x, i) => new { Column = i, Title = x })
                .ToDictionary(x => x.Title, x => x.Column);

            Excel.Range minCell = _sheet.Cells[_titleRow, 1];
            Excel.Range maxCell = _sheet.Cells[objectList.Count + _titleRow, titleColumnDic.Count];
            var rowsCount = maxCell.Row - minCell.Row + _titleRow;
            var columnsCount = maxCell.Column - minCell.Column + 1;
            var data = new object[rowsCount, columnsCount];

            titleColumnDic.OrderBy(x => x.Value).ForEach(x =>
            {
                var row = 0;
                var column = x.Value;
                var title = x.Key;
                data[row, column] = title;
                _cellDataList.Add(new CellData
                {
                    Cell = _sheet.Cells[_titleRow, column + 1],
                    Key = new JsonTitle(title),
                    Type = DataType.Title,
                });
            });

            objectList.ForEach(x =>
            {
                var row = x.Index + 1;
                x.Object.Keys.ForEach(key =>
                {
                    var column = titleColumnDic[key];
                    var jsonToken = x.Object.GetJsonToken(key);
                    if (jsonToken != null)
                    {
                        data[row, column] = jsonToken.ToValue();
                        _cellDataList.Add(new CellData
                        {
                            Index = x.Index,
                            Cell = _sheet.Cells[row + _titleRow, column + 1],
                            Value = jsonToken,
                            Type = DataType.Value,
                        });
                    }
                });
            });

            var range = _sheet.get_Range(minCell.Address, maxCell.Address);
            range.Value2 = data;
        }

        public void Spread(Excel.Range cell)
        {
            cell.Value2 = "[array]";
        }

        public bool OnDoubleClick(Excel.Workbook book, Excel.Range target)
        {
            if (_cellDataList.Empty(x => x.Cell.Address == target.Address))
            {
                return false;
            }

            return true;

            var cellData = _cellDataList.First(x => x.Cell.Address == target.Address);
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
                    _cellDataList = _cellDataList.Where(x => !childs.Contains(x.Key)).ToList();
                    childs.ForEach(x => cellData.Key.RemoveChildTitle(x));

                    _sheet.Range[left, right].EntireColumn.Delete();
                }
                else
                {
                    var titles = _cellDataList.Where(x => x.Cell.Column == cellData.Cell.Column)
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
                    _cellDataList.AddRange(titleCellDatas);
                    //FillCellData(_sheet, _token, _cellDatas);
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
            var cellData = _cellDataList.FirstOrDefault(x => x.Cell.Address == target.Address);

            cellData?.Value.OnChangeValue(target);
        }

        private void SetNamedRange(Excel.Worksheet sheet, IEnumerable<CellData> titles)
        {
            titles.ForEach(x =>
            {
                var jsonTitle = (JsonTitle)x.Key;
                var columnText = x.Cell.Address.Split('$').FirstOrDefault(e => e.Length > 0);
                sheet.Names.Add(jsonTitle.Title, sheet.Range[$"{columnText}:{columnText}"]);
            });
        }
    }
}
