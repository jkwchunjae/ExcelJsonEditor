using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelJsonEditorAddin.Theme;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelJsonEditorAddin
{
    public partial class ExcelJsonEditorRibbon
    {
        public event EventHandler<ThemeType> ChangeTheme;

        private Dictionary<ThemeType, RibbonCheckBox> _themeCheckBoxDic
            = new Dictionary<ThemeType, RibbonCheckBox>();

        public void SetCheckBox(ThemeType themeType)
        {
            _themeCheckBoxDic.ToList()
                .ForEach(x => x.Value.Checked = themeType == x.Key);
        }

        private void ExcelJsonEditorRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _themeCheckBoxDic[ThemeType.White] = ThemeWhiteCheckbox;
            _themeCheckBoxDic[ThemeType.Dark] = ThemeDarkCheckbox;
        }

        private void ThemeWhiteCheckbox_Click(object sender, RibbonControlEventArgs e)
        {
            if (ThemeWhiteCheckbox.Checked == false)
            {
                ThemeWhiteCheckbox.Checked = true;
            }
            else
            {
                ChangeTheme?.Invoke(sender, ThemeType.White);
                ThemeDarkCheckbox.Checked = false;
            }
        }

        private void ThemeDarkCheckbox_Click(object sender, RibbonControlEventArgs e)
        {
            if (ThemeDarkCheckbox.Checked == false)
            {
                ThemeDarkCheckbox.Checked = true;
            }
            else
            {
                ChangeTheme?.Invoke(sender, ThemeType.Dark);
                ThemeWhiteCheckbox.Checked = false;
            }
        }

        private void ThemeCheckBox_Click(RibbonCheckBox sender)
        {
            if (sender.Checked)
            {
                var themeType = _themeCheckBoxDic.First(x => x.Value == sender).Key;
                ChangeTheme?.Invoke(sender, themeType);
                SetCheckBox(themeType);
            }
        }
    }
}
