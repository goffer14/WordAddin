using BackgroundWorkerDemo;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

using System.Windows.Forms;
using eDocs_Editor.Properties;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Threading;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MyRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace eDocs_Editor
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        public Office.IRibbonUI ribbon;
        public static AlertForm  alert;
        public static setStylesToDoc setStyle_frm;
        public static monitoringFrm monitoring_Frm;
        public static LoepCreator multiLoep;
        public static AuthenticateForm Authenticate_frm = null;
        public MyRibbon()
        {
        }
        public Word.Document Doc;
        public static string DocLoction;
        public static bool firstClick = true;
        public static string AutoRevString = "";
        public static string AutoDateString = "";
        object Miss = System.Reflection.Missing.Value;

        public bool unsecure_for_edit(Word.Document DocToUnSecure)
        {
            object password = settings.doc_password;
            try
            {
                DocToUnSecure.Unprotect(ref password);
            }
            catch
            {
                return false;
            }
            return true;
        }
        public bool unsecure_for_edit()
        {
            object password = settings.doc_password;
            try
            {
                Doc.Unprotect(ref password);
            }
            catch
            {
                return false;
            }
            return true;
        }
        public stdole.IPictureDisp GetImage(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "CreateEdoc_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.new_doc);
                    }
                case "finish_page_rivision_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.procces_doc);
                    }
                case "setStyels_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.procces_doc);
                    }
                case "PageRevision_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.edit_rev);
                    }
                case "Edit_doc_Template_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.edit_template);
                    }
                case "ProcessListOfE_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.list_of_effctive); 
                    }
                case "Edit_list_Template_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.edit_listof);
                    }
                case "Header1ToTop_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.reset_header);
                    }
                case "Loep_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.reset_header);
                    }
                case "sameAsPrevious_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.free_text_Img);
                    }
                case "toggleButton_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.edit_rev);
                    }
                case "Change_style_button":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.reset_header);
                    }
                case "ProcessTOC_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.reset_header);
                    }
                case "ExportChanges_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.change_rivision);
                    }
                case "removeUnWantedStyles":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.change_rivision);
                    }
            }
            return null;
        }
        public bool isActiveAddin()
        {
            try { Doc.Fields.Locked = 0; } catch { };
            if (Settings.Default.is_active_auth == "false")
            {
                DialogResult dialogResult = MessageBox.Show("License Required, Enter License?", "License Required", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    AuthenticateForm authFrm = new AuthenticateForm();
                    authFrm.Show();
                }
                return false;
            }
            return true;
        }
        static void menu_Click(object sender, EventArgs e)
        {
            MessageBox.Show(((MenuItem)sender).Text);
        }
        internal class PictureConverter : AxHost
        {
            private PictureConverter() : base(String.Empty) { }

            static public stdole.IPictureDisp ImageToPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }
        public void OnListOf(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            if (settings.monitorDoc)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
            /**
            DialogResult dialogResult = MessageBox.Show("Create list of effective pages in A4 format?", "If A4 format click yes, if A5 format click no", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
                CreateLOEP(4);
            if (dialogResult == DialogResult.No)
                CreateLOEP(5);
            **/

            DialogResult dialogResult = MessageBox.Show("Create list of effective pages?", "To Start click yes, To cancel click no", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
                CreateLOEP(5);
            if (dialogResult == DialogResult.No)
                return;

        }
        private void CreateLOEP(int pageSize)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (!settings.check_if_edoc(Doc))
            {
                MessageBox.Show("eDoc Only");
                return;
            }
            settings.trackChange(Doc, false);
            alert = new AlertForm(Doc, 4, null, pageSize);
            alert.Show();
        }
        public void OnTOC(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            if (settings.monitorDoc)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
                Doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (!settings.check_if_edoc(Doc))
                {
                    MessageBox.Show("eDoc Only");
                    return;
                }
                settings.trackChange(Doc, false);
                alert = new AlertForm(Doc, 8, null, 1);
                alert.Show();
        }
        public void insertSectionPageText(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            insertText("page", Doc.Application.Selection.Range);
        }
        public void insertSectionPageTemplateText(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            string pageCode = "X-P1";
            try{
                pageCode = Doc.Variables["pageTemplate"].Value;
            }
            catch{}
            insertText(pageCode, Doc.Application.Selection.Range);
        }
        public void insertSectionRevText(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            insertText("rev", Doc.Application.Selection.Range);
        }
        public void insertSectionDateText(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            insertText("date", Doc.Application.Selection.Range);
        }
        public void insertSectionIssueText(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            insertText("issue", Doc.Application.Selection.Range);
        }
        public void insertSectionEffictiveText(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            insertText("effective", Doc.Application.Selection.Range);
        }
        public void insertSectionText1Text(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            insertText("text1", Doc.Application.Selection.Range);
        }
        public void insertSectionText2Text(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            insertText("text2", Doc.Application.Selection.Range);
        }
        public void insertSectionText3Text(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            insertText("text3", Doc.Application.Selection.Range);
        }
        public void insertSectionText4Text(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            insertText("text4", Doc.Application.Selection.Range);
        }
        public void insertText(string filedName,Word.Range rng)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (!settings.check_if_edoc(Doc))
            {
                MessageBox.Show("eDoc Only");
                return;
            }
            settings.trackChange(Doc, false);
            DocSettings DS = new DocSettings(Doc,"WTF");
            DS.InesrtRevDatatoAllHeadingCells(filedName,rng);
            DS.UpDateFields();
        }
        public void seteDocsStyels(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            if (setStyle_frm != null)
            {
                setStyle_frm.Close();
                setStyle_frm.Dispose();
                setStyle_frm = null;
            }
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            settings.trackChange(Doc, false);
            setStyle_frm = new setStylesToDoc(Doc);
            setStyle_frm.Show();
        }
        public void openLOEP(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (multiLoep != null)
            {
                multiLoep.Close();
                multiLoep.Dispose();
                multiLoep = null;
            }
            settings.trackChange(Doc, false); 
             multiLoep = new LoepCreator(Doc);
            multiLoep.Show();
        }
        public static void changeStyles(Word.Document Doc)
        {
            alert = new AlertForm(Doc, 5, null,0);
            alert.Show();
        }
        public static void moveToNewVersion(Word.Document Doc)
        {
            alert = new AlertForm(Doc, 9, null, 0);
            alert.Show();
        }
        public void OnExportChagnes(Office.IRibbonControl control)
        {

            if (settings.monitorDoc)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (!settings.check_if_edoc(Doc))
            {
                MessageBox.Show("eDoc Only");
                return;
            }

            alert = new AlertForm(Doc, 3,null, 0);
            alert.Show();
        }
        public void removeUnwantedStyles(Office.IRibbonControl control)
        {
            if (settings.monitorDoc)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (!settings.check_if_edoc(Doc))
            {
                MessageBox.Show("eDoc Only");
                return;
            }
            int totalStyles = Doc.Styles.Count;
            DialogResult dialogResult = MessageBox.Show("Remove unused styles? - styels count: " + totalStyles, "Continue?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                Thread thread = new Thread(removeStyles);
                thread.Start();
            }
        }
        private void removeStyles()
        {
            int deletedStyles = 0;
            int numOfStylesDone = 0;
            foreach (Word.Style mStyle in Doc.Styles)
            {
                if (!mStyle.BuiltIn)
                {
                    Doc.Content.Find.ClearFormatting();
                    Doc.Content.Find.set_Style(mStyle);
                    if (!Doc.Content.Find.Execute())
                    {
                        mStyle.Delete();
                        deletedStyles++;
                    }

                }
                numOfStylesDone++;
                System.Diagnostics.Debug.WriteLine("numOfStylesDone: " + numOfStylesDone);
            }
            MessageBox.Show("Deleted Styles - " + deletedStyles);
        }
        public void autoRev(Office.IRibbonControl control, bool Pressed)
        {
            if (Pressed)
            {
                if (!isActiveAddin())
                    return;
                Doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (!settings.check_if_edoc(Doc))
                {
                    updateMonitorRibbon();
                    MessageBox.Show("eDoc Only");
                    return;
                }
                settings.monitorDoc = true;
                DateTime ts = DateTime.Now;
                settings.dateToMonitor = new DateTime(ts.Year, ts.Month, ts.Day, ts.Hour, ts.Minute, 0);
                Doc.TrackFormatting = true;
                Doc.TrackMoves = true;
                Doc.TrackRevisions = true;
                /**

                if (monitoring_Frm != null)
                {
                    monitoring_Frm.Close();
                    monitoring_Frm.Dispose();
                    monitoring_Frm = null;
                }
                monitoring_Frm = new monitoringFrm(Doc);
                monitoring_Frm.Show();
    **/
                updateMonitorRibbon();
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Do Not Cancel Monitoring - or you will lose all changes.To continue: Press - Yes - and then press Process to save changes ", "Attention !!", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                    updateMonitorRibbon();
                else if (dialogResult == DialogResult.No)
                {
                    settings.monitorDoc = false;
                    updateMonitorRibbon();
                }
                
            }

        }
        public void updateMonitorRibbon()
        {
            ribbon.InvalidateControl("toggleButton_ribbon");
            ribbon.InvalidateControl("rev_cbo");
            ribbon.InvalidateControl("date_cbo");
        }
        public void updateMonitorRibbonFrm()
        {
            ribbon.InvalidateControl("rev_cbo");
            ribbon.InvalidateControl("date_cbo");
        }
        public bool GetEnable(Office.IRibbonControl control)
        {
            return settings.monitorDoc;
        }
        public string GetMonitorText(Office.IRibbonControl control)
        {
            if (settings.monitorDoc)
                return "Stop Monitoring";
            return "Start Monitoring";
        }

        public void ocCurrentRev(Office.IRibbonControl control, string text)
        {
            AutoRevString = text;
        }
        public void ocCurrentDate(Office.IRibbonControl control, string text)
        {
            AutoDateString = text;
            ribbon.InvalidateControl("date_cbo");
        }
        public string onGetEbCurrentDate(Office.IRibbonControl control)
        {
            return AutoDateString;
        }
        public string onGetEbCurrentRev(Office.IRibbonControl control)
        {
            return AutoRevString;
        }
        public void PageRevisionForm(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            if (settings.monitorDoc)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
            if (settings.page_rev_frm != null)
            {
                settings.page_rev_frm.Close();
                settings.page_rev_frm.Dispose();
                settings.page_rev_frm = null;
            }

            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            DocSettings DS = new DocSettings(Doc);

            if (settings.check_if_edoc(Doc))
            {
                Doc.Application.ScreenUpdating = false;
                settings.trackChange(Doc, false);
                settings.page_rev_frm = new PageRevisionFrm(Doc);
                if (settings.last_rev != "")
                    settings.page_rev_frm.page_rev.Text = settings.last_rev;
                if (settings.last_date == "")
                    settings.last_date = DateTime.Now.Date.ToString("d", DateTimeFormatInfo.InvariantInfo);
                settings.page_rev_frm.date_rev.Text = settings.last_date;

                System.Drawing.Point cursor_pos = new System.Drawing.Point(Cursor.Position.X -50, Cursor.Position.Y-10);
                settings.page_rev_frm.Location = cursor_pos;
                Doc.Application.ScreenUpdating = true;
                settings.trackChange(Doc, true);
                settings.page_rev_frm.Show();
            }
            else
            {
                MessageBox.Show("eDoc Only");
            }
        }
        public void makeAllSameAsPrevious(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            if (settings.monitorDoc)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (settings.check_if_edoc(Doc))
            {
                settings.trackChange(Doc, false);
                alert = new AlertForm(Doc, 6,null, 0);
                alert.Show();
            }
            else
            {
                MessageBox.Show("eDoc Only");
            }
        }
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("eDocs_Editor.MyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            AutoDateString = DateTime.Now.Date.ToString("d", DateTimeFormatInfo.InvariantInfo);
            this.ribbon = ribbonUI;
        }
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
