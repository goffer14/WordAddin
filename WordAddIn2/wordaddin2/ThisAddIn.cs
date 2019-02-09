using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Threading;
using Microsoft.Office.Tools.Word;
using eDocs_Editor.Properties;

namespace eDocs_Editor
{
    public partial class ThisAddIn
    {
        public MyRibbon m_Ribbon;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                if (Properties.Settings.Default.UpgradeRequired)
                {
                    Properties.Settings.Default.Upgrade();
                    Properties.Settings.Default.UpgradeRequired = false;
                    Properties.Settings.Default.Save();
                }
                Thread oThread = new Thread(new ThreadStart(settings.check_if_vaild_onStartUp));
                oThread.Start();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        void Application_DocumentOpen(Word.Document document)
        {
            try
            {
                if (document.Fields.Locked == 0)
                    System.Diagnostics.Debug.WriteLine("Application_DocumentOpen - Doc was UNLOCK");
                else
                {
                    document.Fields.Locked = 0;
                    System.Diagnostics.Debug.WriteLine("Application_DocumentOpen - Doc now UNLOCK");
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Application_DocumentOpen Error - " + e.Message);
            }

            try
            {
                string pageCode = document.Variables["pageTemplate"].Value;
            }
            catch (Exception e)
            {
                document.Variables["pageTemplate"].Value = "X-P-1";
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        void App_BeforeSaveDocument(Word.Document document, ref bool saveAsUI, ref bool cancel)
        {
            try {
                if (document.Fields.Locked == -1)
                    System.Diagnostics.Debug.WriteLine("App_BeforeSaveDocument - Doc was LOCK");
                else
                {
                    document.Fields.Locked = -1;
                    System.Diagnostics.Debug.WriteLine("App_BeforeSaveDocument - Doc now LOCK");
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("App_BeforeSaveDocument Error - " + e.Message);
            }
        }

        void App_DocumentBeforePrint(Word.Document document, ref bool cancel)
        {
            try
            {
                if (document.Fields.Locked == -1)
                    System.Diagnostics.Debug.WriteLine("App_DocumentBeforePrint - Doc was LOCK");
                else
                {
                    document.Fields.Locked = -1;
                    System.Diagnostics.Debug.WriteLine("App_DocumentBeforePrint - Doc now LOCK");
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("App_DocumentBeforePrint Error - " + e.Message);
            }
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            m_Ribbon = new MyRibbon();
            return m_Ribbon;
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(App_BeforeSaveDocument);
            this.Application.DocumentBeforePrint += new Word.ApplicationEvents4_DocumentBeforePrintEventHandler(App_DocumentBeforePrint);
        }

        #endregion
    }
}
