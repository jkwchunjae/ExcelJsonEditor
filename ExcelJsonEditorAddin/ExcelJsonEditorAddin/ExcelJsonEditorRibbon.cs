using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelJsonEditorAddin.Config;
using ExcelJsonEditorAddin.Theme;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;

namespace ExcelJsonEditorAddin
{
    public partial class ExcelJsonEditorRibbon
    {
        public event EventHandler<Settings> ChangeSettings;

        private Settings _settings;
        private Dictionary<ThemeType, RibbonCheckBox> _themeCheckBoxDic
            = new Dictionary<ThemeType, RibbonCheckBox>();

        public void Initialize(Settings settings)
        {
            _settings = JsonConvert.DeserializeObject<Settings>(JsonConvert.SerializeObject(settings));
            SetThemeDefaultCheckBox.Checked = settings.SetDefaultExcelTheme;
            SetThemeCheckBox(settings.Theme);
        }

        private void SetThemeCheckBox(ThemeType themeType)
        {
            _themeCheckBoxDic.ToList()
                .ForEach(x => x.Value.Checked = themeType == x.Key);
        }

        private void ExcelJsonEditorRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _themeCheckBoxDic[ThemeType.White] = ThemeWhiteCheckbox;
            _themeCheckBoxDic[ThemeType.Dark] = ThemeDarkCheckbox;
        }

        private void ThemeCheckBox_Click(object sender, RibbonControlEventArgs e)
        {
            var checkbox = (RibbonCheckBox)sender;
            if (checkbox.Checked)
            {
                var themeType = _themeCheckBoxDic.First(x => x.Value == checkbox).Key;
                _settings.Theme = themeType;
                ChangeSettings?.Invoke(sender, _settings);
                SetThemeCheckBox(themeType);
            }
        }

        private void SetThemeDefaultCheckBox_Click(object sender, RibbonControlEventArgs e)
        {
            _settings.SetDefaultExcelTheme = ((RibbonCheckBox)sender).Checked;
            ChangeSettings?.Invoke(sender, _settings);
        }
    }
}
