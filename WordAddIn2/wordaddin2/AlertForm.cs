using System;
using System.ComponentModel;
using System.Windows.Forms;
using WordAddIn2;
using Word = Microsoft.Office.Interop.Word;

namespace BackgroundWorkerDemo
{
    public partial class AlertForm : Form
    {

        #region PROPERTIES

        public Word.Document Doc;
        public BackgroundWorker worker;
        public int WhatToDo;
        public int FromPage;
        string filename;
        #endregion


        public AlertForm(Word.Document doc, int WhatToDo, int FromPage , string filename)
        {
            InitializeComponent();
            if (backgroundWorker1.IsBusy != true)
            {
                this.WhatToDo = WhatToDo;
                this.Doc = doc;
                this.filename = filename;
                this.FromPage = FromPage;
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            worker = sender as BackgroundWorker;
            if(WhatToDo!= 6)
                settings.trackChange(Doc, false);
            settings.alert = this;
            switch (WhatToDo)
            {
                case 1:
                    settings.process_doc(Doc);
                    break;
                case 2:
                    settings.process_doc(Doc); 
                    break;
                case 3:
                    settings.init_ListOfE_New(Doc , filename);
                    break;
                case 4:
                    settings.Header1ToTop(Doc);
                    break;
                case 5:
                    settings.OnCreateEdocFromDoc(Doc,filename);
                    break;
                case 6:
                    settings.SavePagesText(Doc);
                    break;
                case 7:
                    settings.changeStyles(Doc);
                    break;
            }

        }

        // This event handler updates the progress.
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Pass the progress to AlertForm label and progressbar
        }

        // This event handler deals with the results of the background operation.
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            settings.trackChange(Doc, true);
            settings.alert.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            settings.trackChange(Doc, true);
            this.backgroundWorker1.CancelAsync();
            this.Close();
        }

        private void AlertForm_Load(object sender, EventArgs e)
        {

        }



    }
}
