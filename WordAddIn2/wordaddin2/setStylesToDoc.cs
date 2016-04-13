using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordAddIn2.Properties;

using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn2
{
    public partial class setStylesToDoc : Form
    {
        private Word.Document Doc;
        private string fileName;
        public static string[] pageNameString = null;
        public setStylesToDoc(Word.Document ParentDoc,string fileName)
        {
            Cursor.Current = Cursors.WaitCursor;
            this.Doc = ParentDoc;
            this.fileName = fileName;
            try
            {
                string test1 = Doc.Variables["heading1_name"].Value;
            }
            catch
            {
                Doc.Variables.Add("heading1_name", "Heading 1");
            }
            try
            {
                string test1 = Doc.Variables["heading2_name"].Value;
            }
            catch
            {
                Doc.Variables.Add("heading2_name", "Heading 2");
            }
            try
            {
                string test1 = Doc.Variables["introduction1"].Value;
            }
            catch
            {
                Doc.Variables.Add("introduction1", "Introduction‎ 1");
            }
            try
            {
                string test1 = Doc.Variables["introduction2"].Value;
            }
            catch
            {
                Doc.Variables.Add("introduction2", "Introduction‎ 2");
            }
            InitializeComponent();
        }
        private void btn_ok_Click(object sender, EventArgs e)
        {
            if(H1.SelectedIndex != H2.SelectedIndex && H1.SelectedIndex!=I1.SelectedIndex && H1.SelectedIndex != I2.SelectedIndex)
                if (H2.SelectedIndex != I1.SelectedIndex && H2.SelectedIndex != I2.SelectedIndex)
                    if (I1.SelectedIndex != I2.SelectedIndex)
                    {
                        Doc.Variables["heading1_name"].Value = H1.Items[H1.SelectedIndex].ToString();
                        Doc.Variables["heading2_name"].Value = H2.Items[H2.SelectedIndex].ToString();
                        Doc.Variables["introduction1"].Value = I1.Items[I1.SelectedIndex].ToString();
                        Doc.Variables["introduction2"].Value = I2.Items[I2.SelectedIndex].ToString();
                        if(fileName!=null)
                            MyRibbon.CreateEdocFromDoc(Doc, fileName);
                        else
                            MyRibbon.changeStyles(Doc);
                        this.Close();
                        this.Dispose();
                        return;
                    }
            MessageBox.Show("Value mistake - All styles must be different");
        }

        private void setStylesToDoc_Load(object sender, EventArgs e)
        {
            int selectedH1 = 0;
            int selectedH2 = 0;
            int selectedI1 = 0;
            int selectedI2 = 0;
            for (int i=1;i<=Doc.Styles.Count;i++)
            {
                H1.Items.Add(Doc.Styles[i].NameLocal);
                if (Doc.Variables["heading1_name"].Value.Equals(Doc.Styles[i].NameLocal))
                    selectedH1 = i - 1;
                H2.Items.Add(Doc.Styles[i].NameLocal);
                if (Doc.Variables["heading2_name"].Value.Equals(Doc.Styles[i].NameLocal))
                    selectedH2 = i - 1;
                I1.Items.Add(Doc.Styles[i].NameLocal);
                if (Doc.Variables["introduction1"].Value.Equals(Doc.Styles[i].NameLocal))
                    selectedI1 = i - 1;
                I2.Items.Add(Doc.Styles[i].NameLocal);
                if (Doc.Variables["introduction2"].Value.Equals(Doc.Styles[i].NameLocal))
                    selectedI2 = i - 1;
            }

            H1.SelectedIndex = selectedH1;
            H2.SelectedIndex = selectedH2;
            I1.SelectedIndex = selectedI1;
            I2.SelectedIndex = selectedI2;
            Cursor.Current = Cursors.Default;
        }
    }
}
