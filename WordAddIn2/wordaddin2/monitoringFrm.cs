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
    public partial class monitoringFrm : Form
    {
        private Microsoft.Office.Interop.Word.Document Doc;
        DocSettings DS;
        public monitoringFrm(Microsoft.Office.Interop.Word.Document ParentDoc)
        {
            this.Doc = ParentDoc;
            InitializeComponent();
        }

        private void monitoringFrm_Load(object sender, EventArgs e)
        {
            try
            {
                this.date_rev.Text = Doc.Variables["last_date_rivision"].Value;
            }
            catch { }
            try
            {
                this.page_rev.Text = Doc.Variables["last_page_rivision"].Value;
            }
            catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
                MyRibbon.AutoRevString = page_rev.Text;
                MyRibbon.AutoDateString = date_rev.Text;
            this.Close();
            this.Dispose();
        }
    }
}