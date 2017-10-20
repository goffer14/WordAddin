using System;
using System.ComponentModel;
using System.Windows.Forms;
using eDocs_Editor;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace BackgroundWorkerDemo
{
    public partial class AlertForm : Form
    {

        #region PROPERTIES

        public Word.Document Doc;
        public BackgroundWorker worker;
        public int WhatToDo;
        public int pageSize;
        public List<loepDocument> loepDocumentArray;
        #endregion


        public AlertForm(Word.Document doc, int WhatToDo,List<loepDocument> loepDocumentArray,int pageSize)
        {
            InitializeComponent();
            if (backgroundWorker1.IsBusy != true)
            {
                this.WhatToDo = WhatToDo;
                this.pageSize = pageSize;
                this.Doc = doc;
                if(loepDocumentArray!=null)
                    this.loepDocumentArray = new List<loepDocument>(loepDocumentArray);
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            worker = sender as BackgroundWorker;
            if(WhatToDo!= 1)
                settings.trackChange(Doc, false);
            settings.alert = this;
            switch (WhatToDo)
            {
                case 2:
                    settings.ProcessMonitoring(Doc);
                    break;
                case 3:
                    settings.ChangesExport(Doc);
                    break;
                case 4:
                    settings.init_ListOfE_New(Doc, pageSize);
                    break;
                case 5:
                    settings.process_doc(Doc);
                    break;
                case 6:
                    settings.makeAllSameAsPrevious(Doc);
                    break;
                case 7:
                    settings.CreateMultiLOEP(Doc, loepDocumentArray);
                    break;
                case 8:
                    settings.initTOC(Doc);
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
