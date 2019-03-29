namespace ExcelJsonEditorAddin
{
    partial class ExcelJsonEditorRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExcelJsonEditorRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.JsonEditorTab = this.Factory.CreateRibbonTab();
            this.ThemeGroup = this.Factory.CreateRibbonGroup();
            this.SetThemeDefaultCheckBox = this.Factory.CreateRibbonCheckBox();
            this.ThemeWhiteCheckbox = this.Factory.CreateRibbonCheckBox();
            this.ThemeDarkCheckbox = this.Factory.CreateRibbonCheckBox();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.OpenJsonButton = this.Factory.CreateRibbonButton();
            this.JsonEditorTab.SuspendLayout();
            this.ThemeGroup.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // JsonEditorTab
            // 
            this.JsonEditorTab.Groups.Add(this.ThemeGroup);
            this.JsonEditorTab.Groups.Add(this.group1);
            this.JsonEditorTab.Label = "JsonEditor";
            this.JsonEditorTab.Name = "JsonEditorTab";
            // 
            // ThemeGroup
            // 
            this.ThemeGroup.Items.Add(this.SetThemeDefaultCheckBox);
            this.ThemeGroup.Items.Add(this.ThemeWhiteCheckbox);
            this.ThemeGroup.Items.Add(this.ThemeDarkCheckbox);
            this.ThemeGroup.Label = "Theme";
            this.ThemeGroup.Name = "ThemeGroup";
            // 
            // SetThemeDefaultCheckBox
            // 
            this.SetThemeDefaultCheckBox.Label = "SetDefault";
            this.SetThemeDefaultCheckBox.Name = "SetThemeDefaultCheckBox";
            this.SetThemeDefaultCheckBox.ScreenTip = "Set theme normal excel";
            this.SetThemeDefaultCheckBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetThemeDefaultCheckBox_Click);
            // 
            // ThemeWhiteCheckbox
            // 
            this.ThemeWhiteCheckbox.Label = "White";
            this.ThemeWhiteCheckbox.Name = "ThemeWhiteCheckbox";
            this.ThemeWhiteCheckbox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ThemeCheckBox_Click);
            // 
            // ThemeDarkCheckbox
            // 
            this.ThemeDarkCheckbox.Label = "Dark";
            this.ThemeDarkCheckbox.Name = "ThemeDarkCheckbox";
            this.ThemeDarkCheckbox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ThemeCheckBox_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.OpenJsonButton);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // OpenJsonButton
            // 
            this.OpenJsonButton.Label = "OpenJson";
            this.OpenJsonButton.Name = "OpenJsonButton";
            // 
            // ExcelJsonEditorRibbon
            // 
            this.Name = "ExcelJsonEditorRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.JsonEditorTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExcelJsonEditorRibbon_Load);
            this.JsonEditorTab.ResumeLayout(false);
            this.JsonEditorTab.PerformLayout();
            this.ThemeGroup.ResumeLayout(false);
            this.ThemeGroup.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab JsonEditorTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ThemeGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ThemeWhiteCheckbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ThemeDarkCheckbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenJsonButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox SetThemeDefaultCheckBox;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelJsonEditorRibbon ExcelJsonEditorRibbon
        {
            get { return this.GetRibbon<ExcelJsonEditorRibbon>(); }
        }
    }
}
