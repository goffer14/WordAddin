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

namespace WordAddIn2
{
    public partial class FreeTextfrm : Form
    {
        private Word.Document Doc;
        public static string[] pageNameString = null;
        public FreeTextfrm(Word.Document ParentDoc)
        {
            this.Doc = ParentDoc;
            InitializeComponent();
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            this.Close();
            this.Dispose();
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            DocSettings DS = new DocSettings(Doc);
            settings.trackChange(Doc, false);
            if (radioButton1.Checked)
                DS.ChangeText(pageNameString,0, pageNameString.Length-1, this.text1.Text, this.text2.Text, this.text3.Text, this.text4.Text);
            else if(radioButton2.Checked|| radioButton3.Checked)
            {

                int fromPage = 0;
                int toPage = 0;
                if (radioButton2.Checked)
                {
                    fromPage = FromPagebox.SelectedIndex;
                    toPage = ToPagebox.SelectedIndex;
                }
                else
                {
                    fromPage = FromChapterbox.SelectedIndex;
                    toPage = ToChapterbox.SelectedIndex;
                }
                if (fromPage > toPage)
                {
                    MessageBox.Show("Value mistake");
                    settings.trackChange(Doc, true);
                    return;
                }
                DS.ChangeText(pageNameString, fromPage, toPage, this.text1.Text, this.text2.Text, this.text3.Text, this.text4.Text);
            }
            else
            {
                string str;
                for (int i = 0; i < pageNameString.Count(); i++)
                {
                    try {
                        str = Doc.Variables["edocs_PAGE" + pageNameString[i] + "_rev"].Value;
                        str = Doc.Variables["edocs_PAGE" + pageNameString[i] + "_date"].Value;
                    }
                    catch
                    {
                        DS.insertTextToData(pageNameString[i], this.text1.Text, this.text2.Text, this.text3.Text, this.text4.Text);
                        continue;
                    }
                }
                
            }
            settings.last_text1 = this.text1.Text;
            settings.last_text2 = this.text2.Text;
            settings.last_text3 = this.text3.Text;
            settings.last_text4 = this.text4.Text;

            DS.UpDateFields();
            Cursor.Current = Cursors.Default;
            this.Close();
            this.Dispose();
             
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton2.Checked)
            {
                this.Size = new Size(367, 365);
                label4.Text = "From Page:";
                label5.Text = "To Page:";
                FromChapterbox.Visible = false;
                ToChapterbox.Visible = false;
                FromPagebox.Visible = true;
                ToPagebox.Visible = true;
                
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                this.Size = new Size(367, 295);
            }
        }

        private void PageRevisionFrm_Load(object sender, EventArgs e)
        {
            radioButton1.Checked = true;
            DocSettings DS = new DocSettings(Doc);
            settings.trackChange(Doc, false);
            int DocPageNumber = DS.GetPageNumber(Doc);
            int z = DS.GetFirstHeader()[1]-1;
            
            pageNameString = new string[DocPageNumber - z];
            for (int i = 1 +z; i <= DocPageNumber; i++)
            {
                FromPagebox.Items.Add(i);
                ToPagebox.Items.Add(i);
                string pageName = "";
                try
                {
                    pageName = Doc.Variables["edocs_PAGE" + (i) + "_page"].Value;
                    pageNameString[i-1-z] = pageName;
                }
                catch
                {
                    MessageBox.Show("Doc Have To Be Procees");
                    settings.trackChange(Doc, true);
                    this.Close();
                    this.Dispose();
                    return;
                }
                pageName = pageName.Replace("\r\n", "").Replace("\r", "").Replace("\n", "").Replace("\a", "");
                FromChapterbox.Items.Add(pageName);
                ToChapterbox.Items.Add(pageName);
            }
            if(DS.Heading_pos1.text1_pos==-1)
                text1.Enabled = false;
            if (DS.Heading_pos1.text2_pos == -1)
                text2.Enabled = false;
            if (DS.Heading_pos1.text3_pos == -1)
                text3.Enabled = false;
            if (DS.Heading_pos1.text4_pos == -1)
                text4.Enabled = false;
            settings.trackChange(Doc, true);
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                this.Size = new Size(367, 365);
                label4.Text = "From Chapter:";
                label5.Text = "To Chapter:";
                FromChapterbox.Visible = true;
                ToChapterbox.Visible = true;
                FromPagebox.Visible = false;
                ToPagebox.Visible = false;

            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                this.Size = new Size(367, 295);
            }
        }
    }
}
