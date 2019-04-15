using ExcelJsonEditorAddin.Theme;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonEditorAddin.Extensions
{
    public static class StyleExtensions
    {
        public static Excel.Style SetDefaultStyle(this Excel.Style style, ThemeType themeType)
        {
            var ignoreList = new List<Excel.XlBordersIndex>
            {
                Excel.XlBordersIndex.xlDiagonalDown,
                Excel.XlBordersIndex.xlDiagonalUp,
            };

            switch (themeType)
            {
                case ThemeType.White:
                    style.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
                    style.Font.Color = Color.Black;
                    style.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    break;
                case ThemeType.Dark:
                    style.Interior.Color = Color.FromArgb(30, 30, 30);
                    style.Font.Color = Color.FromArgb(220, 220, 220);

                    style.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    style.Borders.Color = Color.FromArgb(80, 80, 80);
                    style.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    ignoreList.ForEach(index =>
                    {
                        style.Borders[index].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    });
                    break;
            }
            return style;
        }

        public static Excel.Style SetNumberStyle(this Excel.Style style, ThemeType themeType)
        {
            style = style.SetDefaultStyle(themeType);

            style.Font.Name = "Consolas";

            switch (themeType)
            {
                case ThemeType.White:
                    break;
                case ThemeType.Dark:
                    style.Font.Color = Color.FromArgb(181, 206, 168);
                    break;
            }
            return style;
        }

        public static Excel.Style SetStringStyle(this Excel.Style style, ThemeType themeType)
        {
            style = style.SetDefaultStyle(themeType);

            style.NumberFormat = "@";

            switch (themeType)
            {
                case ThemeType.White:
                    break;
                case ThemeType.Dark:
                    break;
            }
            return style;
        }
    }
}
