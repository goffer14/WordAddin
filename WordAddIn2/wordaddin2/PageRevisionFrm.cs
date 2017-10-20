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
    public partial class PageRevisionFrm : Form
    {
        private Word.Document Doc;
        private Int32 firstClickItemIndex = 0;
        DocSettings DS;
        public PageRevisionFrm(Word.Document ParentDoc)
        {
            this.Doc = ParentDoc;
            InitializeComponent();
        }

        private void PageRevisionFrm_Load(object sender, EventArgs e)
        {
            DS = new DocSettings(Doc);
            settings.trackChange(Doc, false);
            int DocPageNumber = DS.GetPageNumber(Doc);
            string pageRev = "";
            string eDocPage = "";
            string pageDate = "";
            string pageIssue = "";
            string pageEffectiveDate = "";
            string text1 = "";
            string text2 = "";
            string text3 = "";
            string text4 = "";
            firstClickItemIndex = 0;
            try
            {
                this.date_rev.Text = Doc.Variables["last_date_rivision"].Value;
            }
            catch  { }
            try
            {
                this.page_rev.Text = Doc.Variables["last_page_rivision"].Value;
            }
            catch{  }
            try
            {
                this.issueText.Text = Doc.Variables["last_issue"].Value;
            }
            catch { }
            try
            {
                this.effectiveText.Text = Doc.Variables["last_effective_date"].Value;
            }
            catch { }
            try
            {
                this.text1_text.Text = Doc.Variables["last_text1_value"].Value;
            }
            catch { }
            try
            {
                this.text2_text.Text = Doc.Variables["last_text2_value"].Value;
            }
            catch { }

            try
            {
                this.text3_text.Text = Doc.Variables["last_text3_value"].Value;
            }
            catch { }
            try
            {
                this.text4_text.Text = Doc.Variables["last_text4_value"].Value;
            }
            catch { }
            for (int i = 1; i <= DocPageNumber; i++)
            {
                Cursor.Current = Cursors.WaitCursor;
                pageRev = getVarString("rev", i);
                eDocPage = getVarString("page", i);
                pageDate = getVarString("date", i);
                pageIssue = getVarString("issue", i);
                pageEffectiveDate = getVarString("effective", i);
                text1 = getVarString("text1", i);
                text2 = getVarString("text2", i);
                text3 = getVarString("text3", i);
                text4 = getVarString("text4", i);
                string[] row = { i.ToString(), eDocPage, pageRev, pageDate, pageIssue, pageEffectiveDate, text1, text2, text3, text4 };
                var listViewItem = new ListViewItem(row);
                pageView.Items.Add(listViewItem);
            }
            Cursor.Current = Cursors.Default;
            settings.trackChange(Doc, true);
        }
        public string getVarString(string varName, int page)
        {
            string endVar;
            if (varName == "page")
                try
                {
                    endVar = Doc.Variables["edocs_Page" + page + "_page"].Value;
                    return endVar.Replace("\r\n", "").Replace("\r", "").Replace("\n", "").Replace("\a", "");
                }
                catch { return "-"; }
            try
            {
                endVar = Doc.Variables["edocs_Page" + Doc.Variables["edocs_Page" + page + "_page"].Value + "_" + varName].Value;
                return endVar.Replace("\r\n", "").Replace("\r", "").Replace("\n", "").Replace("\a", "");
            }
            catch { return "-"; }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            int DocPageNumber = DS.GetPageNumber(Doc);
            settings.trackChange(Doc, false);
            string[] vars = new string[9];
            vars[1] = addVar("last_page_rivision", this.page_rev);
            vars[2] = addVar("last_date_rivision", this.date_rev);
            vars[3] = addVar("last_issue", this.issueText);
            vars[4] = addVar("last_effective_date", this.effectiveText);
            vars[5] = addVar("last_text1_value", this.text1_text);
            vars[6] = addVar("last_text2_value", this.text2_text);
            vars[7] = addVar("last_text3_value", this.text3_text);
            vars[8] = addVar("last_text4_value", this.text4_text);
            if (checkBox1.Checked)
            {
                DialogResult dialogResult = MessageBox.Show("Be aware this action will override all existing Dates and Rivisions data in all Pages. are you sure?", "Update all Doc?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    string edoc_page_text = "";
                    for (int i = 0; i < pageView.Items.Count; i++)
                    {
                        Cursor.Current = Cursors.WaitCursor;
                        edoc_page_text = "edocs_Page" + pageView.CheckedItems[i].SubItems[0].Text + "_page";
                        try {
                            vars[0] = Doc.Variables[edoc_page_text].Value;
                            saveChanges(vars,i);
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    Cursor.Current = Cursors.Default;
                }

            }
            else
            {
                string edoc_page_text = "";
                for (int i = 0; i < pageView.CheckedItems.Count; i++)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    edoc_page_text = "edocs_Page" + pageView.CheckedItems[i].SubItems[0].Text + "_page";
                    try
                    {
                        vars[0] = Doc.Variables[edoc_page_text].Value;
                        saveChanges(vars,i);
                    }
                    catch
                    {
                        continue;
                    }
                }
                Cursor.Current = Cursors.Default;

            }
            Cursor.Current = Cursors.Default;
            DS.UpDateFields();
        }
        private void saveChanges(string[] vars,int i)
        {
            DS.insertRev_Rdate(vars[0], vars[1], vars[2], vars[3], vars[4], vars[5], vars[6], vars[7], vars[8]);
            for (int z = 1; z <= 8; z++)
            {
                if(vars[z]!=null)
                    pageView.CheckedItems[i].SubItems[z+1].Text = vars[z];
            }

        }
        private string addVar(string varText,TextBox textBox)
        {
            if (textBox.Enabled == false)
                return null;
            try
            {
                Doc.Variables[varText].Value = textBox.Text;
            }
            catch
            {
                Doc.Variables.Add(varText, textBox.Text);
            }
            return textBox.Text;
        }
        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                CheckAllItems(true);
            else
                CheckAllItems(false);
        }
        
        public void CheckAllItems(bool check)
        {
            pageView.Items.OfType<ListViewItem>().ToList().ForEach(item => item.Checked = check);
        }
        private void pageView_ItemCheck(Object sender, ItemCheckedEventArgs e)
        {
            checkBox1.Checked = false;
        }

        private void pageView_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void pageView_MouseClick(object sender, MouseEventArgs e)
        {
            try {
                var info = pageView.HitTest(e.X, e.Y);
                var row = info.Item.Index;
                var col = info.Item.SubItems.IndexOf(info.SubItem);
                var value = info.Item.SubItems[col].Text;
                Clipboard.SetText(value);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex);
            }
        }
        public Boolean numInRange(Int32 num, Int32 first, Int32 last)
        {
            return first <= num && num <= last;
        }

        private void pageView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.A && e.Control)
            {
                foreach (ListViewItem lvItem in pageView.Items)
                    lvItem.Selected = true;
            }
        }

        private void pageView_MouseDown(object sender, MouseEventArgs e)
        {
            if (MouseButtons == MouseButtons.Left && pageView.HitTest(e.Location).Item != null)
                firstClickItemIndex = pageView.HitTest(e.Location).Item.Index;
        }

        private void pageView_MouseMove(object sender, MouseEventArgs e)
        {
            if (MouseButtons == MouseButtons.Left)
            {
                ListViewItem lvItem = pageView.HitTest(e.Location).Item;
                if (lvItem != null)
                {
                    lvItem.Selected = true;
                    if (pageView.SelectedItems.Count > 1)
                    {
                        Int32 firstSelected = pageView.SelectedItems[0].Index;
                        Int32 lastSelected = pageView.SelectedItems[pageView.SelectedItems.Count - 1].Index;
                        foreach (ListViewItem tempLvItem in pageView.Items)
                        {
                            if (numInRange(tempLvItem.Index, firstSelected, lastSelected) && (numInRange(tempLvItem.Index, lvItem.Index, firstClickItemIndex) || numInRange(tempLvItem.Index, firstClickItemIndex, lvItem.Index)))
                            {
                                tempLvItem.Selected = true;
                            }
                            else
                            {
                                tempLvItem.Selected = false;
                            }
                        }
                    }
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void revCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            page_rev.Enabled = revCheckBox.Checked;
        }

        private void dateCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            date_rev.Enabled = dateCheckBox.Checked;
        }

        private void text1CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            text1_text.Enabled = text1CheckBox.Checked;
        }

        private void text2CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            text2_text.Enabled = text2CheckBox.Checked;
        }

        private void text3CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            text3_text.Enabled = text3CheckBox.Checked;
        }

        private void text4CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            text4_text.Enabled = text4CheckBox.Checked;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Did you saved all of the changes?", "Close?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                this.Close();
                this.Dispose();
            }
        }
        private void label1_Click(object sender, EventArgs e)
        {
            CheckAllItems(false);
        }

        private void issueCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            issueText.Enabled = issueCheckBox.Checked;
        }

        private void effectiveCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            effectiveText.Enabled = effectiveCheckBox.Checked;
        }
    }
}
