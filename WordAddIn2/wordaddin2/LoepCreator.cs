using BackgroundWorkerDemo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace eDocs_Editor
{
    public partial class LoepCreator : Form
    {
        private string docPath = "";
        public Microsoft.Office.Interop.Word.Document Doc;
        public static AlertForm alert;
        public LoepCreator(Microsoft.Office.Interop.Word.Document Doc)
        {
            this.Doc = Doc;
            InitializeComponent();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory =
            openFile.Title = "Select Document";
            openFile.FileName = "";
            openFile.Multiselect = true;
            openFile.Filter = "Word Documents (*.doc;*.docx)|*.doc;*.docx";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string[] fileNames = openFile.FileNames;
               
                docPath = openFile.InitialDirectory;
                for (int i = 0;i < fileNames.Length;i++)
                {
                    string[] row = { (pageView.Items.Count + 1).ToString(), System.IO.Path.GetFileName(fileNames[i]), fileNames[i] };
                    var listViewItem = new ListViewItem(row);
                    pageView.Items.Add(listViewItem);
                }
            }
        }

        private void LoepCreator_Load(object sender, EventArgs e)
        {
            docPath = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Please Save current Doc before proceeding, to countinue?", "Create LOEP?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                List<loepDocument> loepDocumentArray = new List<loepDocument>();
                foreach (ListViewItem listItem in pageView.Items)
                {
                    loepDocument item = new loepDocument(listItem.SubItems[1].Text, listItem.SubItems[2].Text);
                    loepDocumentArray.Add(item);
                }
                alert = new AlertForm(Doc, 7, loepDocumentArray, 0);
                alert.Show();
                this.Close();
                this.Dispose();
            }
            else
            {
                this.Close();
                this.Dispose();
            }

        }
    }
}
