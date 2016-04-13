using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn2.Properties;
using System.Threading;

namespace WordAddIn2
{
    public partial class ThisAddIn
    {

        public MyRibbon m_Ribbon;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Thread oThread = new Thread(new ThreadStart(settings.check_if_vaild_copy));
            oThread.Start();
        }

        void Application_DocumentOpen(Word.Document Doc)
        {
           
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
        }
        
        #endregion
    }
}
