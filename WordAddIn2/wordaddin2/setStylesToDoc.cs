using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace eDocs_Editor
{
    public partial class setStylesToDoc : Form
    {
        private Word.Document Doc;
        public static string[] pageNameString = null;
        public setStylesToDoc(Word.Document ParentDoc)
        {
            Cursor.Current = Cursors.WaitCursor;
            this.Doc = ParentDoc;
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
                string test1 = Doc.Variables["introduction2"].Value;
            }
            catch
            {
                Doc.Variables.Add("introduction2", "Empty");
            }
            try
            {
                string test1 = Doc.Variables["appendix2"].Value;
            }
            catch
            {
                Doc.Variables.Add("appendix2", "Empty");
            }
            InitializeComponent();
        }
        private void btn_ok_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                Doc.Variables["processType"].Value = "styles";
                if (H2.SelectedIndex != I2.SelectedIndex && I2.SelectedIndex != A2.SelectedIndex)
                    if (H2.SelectedIndex != A2.SelectedIndex)
                    {
                        Doc.Variables["heading2_name"].Value = H2.Items[H2.SelectedIndex].ToString();
                        if (checkBox1.Checked)
                            Doc.Variables["introduction2"].Value = I2.Items[I2.SelectedIndex].ToString();
                        else
                            Doc.Variables["introduction2"].Value = "Empty";
                        if (checkBox3.Checked)
                            Doc.Variables["appendix2"].Value = A2.Items[A2.SelectedIndex].ToString();
                        else
                            Doc.Variables["appendix2"].Value = "Empty";
                        MyRibbon.changeStyles(Doc);
                        this.Close();
                        this.Dispose();
                        return;
                    }
                MessageBox.Show("Value mistake - All styles must be different");
            }
            else
            {
                Doc.Variables["processType"].Value = "pages";
                MyRibbon.changeStyles(Doc);
                this.Close();
                this.Dispose();
            }
        }

        private void setStylesToDoc_Load(object sender, EventArgs e)
        {
            int selectedH2 = 0;
            int selectedI2 = 1;
            int selectedA2 = 3;
            object[] mylistSource = new object[Doc.Styles.Count];
            // populate source with test data
            for (int i = 1; i <= Doc.Styles.Count; i++)
            {
                mylistSource[i-1] = Doc.Styles[i].NameLocal;
                if (selectedH2 == 0 && Doc.Variables["heading2_name"].Value.Equals(Doc.Styles[i].NameLocal))
                    selectedH2 = i - 1;
                if (selectedI2 == 1 && Doc.Variables["introduction2"].Value.Equals(Doc.Styles[i].NameLocal))
                    selectedI2 = i - 1;
                if (selectedA2 == 3 && Doc.Variables["appendix2"].Value.Equals(Doc.Styles[i].NameLocal))
                    selectedA2 = i - 1;
            }
            H2.Items.AddRange(mylistSource);
            I2.Items.AddRange(mylistSource);
            A2.Items.AddRange(mylistSource);
            if (Doc.Variables["introduction2"].Value.Equals("Empty"))
            {
                checkBox1.Checked = false;
                I2.Enabled = false;
            }
            else
            {
                checkBox1.Checked = true;
                I2.Enabled = true;
            }
       
            if (Doc.Variables["appendix2"].Value.Equals("Empty"))
            {
                checkBox3.Checked = false;
                A2.Enabled = false;
            }
            else
            {
                checkBox3.Checked = true;
                A2.Enabled = true;
            }
            H2.SelectedIndex = selectedH2;
            I2.SelectedIndex = selectedI2;
            A2.SelectedIndex = selectedA2;
            try {string test1 = Doc.Variables["processType"].Value;}
            catch{Doc.Variables.Add("processType", "styles");}
            if (Doc.Variables["processType"].Value.Equals("styles"))
                radioButton1.Checked = true;
            else
                radioButton2.Checked = true;
            Cursor.Current = Cursors.Default;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
                I2.Enabled = checkBox1.Checked;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            A2.Enabled = checkBox3.Checked;
        }

        private void A2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
                I2.Enabled = radioButton1.Checked;
                H2.Enabled = radioButton1.Checked;
                A2.Enabled = radioButton1.Checked;
                checkBox1.Enabled = radioButton1.Checked;
                checkBox3.Enabled = radioButton1.Checked;
                label3.Enabled = radioButton1.Checked;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            I2.Enabled = !radioButton2.Checked;
            H2.Enabled = !radioButton2.Checked;
            A2.Enabled = !radioButton2.Checked;
            checkBox1.Enabled = !radioButton2.Checked;
            checkBox3.Enabled = !radioButton2.Checked;
            label3.Enabled = !radioButton2.Checked;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyRibbon.moveToNewVersion(Doc);
            this.Close();
            this.Dispose();
        }
    }
}
