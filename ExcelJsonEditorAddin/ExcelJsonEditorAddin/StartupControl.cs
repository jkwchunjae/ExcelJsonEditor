using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Utils;

namespace ExcelJsonEditorAddin
{
    public partial class StartupControl : UserControl
    {
        public event EventHandler<string> OpenFiles;

        public StartupControl()
        {
            InitializeComponent();

            InitJsonLoadDialog();
        }

        private void InitJsonLoadDialog()
        {
            JsonLoadDialog.Filter = "Json files (*.json)|*.json";
            JsonLoadDialog.FileOk += JsonLoadDialog_FileOk;
        }

        private void JsonLoadDialog_FileOk(object sender, CancelEventArgs e)
        {
            var dialog = (OpenFileDialog)sender;
            //OpenFiles?.Invoke(sender, dialog.FileNames);
            //MessageBox.Show(dialog.FileName);
        }

        private void OpenDialogButton_Click(object sender, EventArgs e)
        {
            var result = JsonLoadDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                OpenFiles?.Invoke(sender, JsonLoadDialog.FileName);
            }
        }
    }
}
