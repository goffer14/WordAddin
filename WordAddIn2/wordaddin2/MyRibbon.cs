using BackgroundWorkerDemo;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

using System.Windows.Forms;
using WordAddIn2.Properties;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

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


namespace WordAddIn2
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        public Office.IRibbonUI ribbon;
        public static AlertForm  alert;
        public static setStylesToDoc setStyle_frm;
        public MyRibbon()
        {
          
        }
        public Word.Document Doc;
        public static string DocLoction;
        public static bool firstClick = true;
        public static string AutoRevString = "";
        public static string AutoDateString = "";
        object Miss = System.Reflection.Missing.Value;

        public static string DocPath()
        {
             string DocPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            DocPath = DocPath + "\\Global eDocs\\eDocs Add-in\\Templates";
            DocLoction = DocPath;
             return DocPath;

        }
        /**
        public void OnCreateEdoc(Office.IRibbonControl control)
        {
            if (settings.escape)
            {
                MessageBox.Show("Evaluation Ended");
                return;
            }
            if(settings.save_text)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = DocPath();
            openFile.Title = "Select Word Template";
            openFile.FileName = "";
            openFile.Filter = "Word Documents (*.doc;*.docx)|*.doc;*.docx";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Doc = Globals.ThisAddIn.Application.Documents.Add(openFile.FileName);
            
            Doc.Application.Selection.WholeStory();
            Doc.Application.Selection.Delete();
                if (!settings.init_doc(Doc))
                {
                    object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
                    object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
                    object routeDocument = false;
                    Doc.Close(ref saveOption, ref originalFormat, ref routeDocument);
                }
            Doc.Application.ActiveWindow.VerticalPercentScrolled = 0;
            Doc.Application.ActiveWindow.View.ShowFieldCodes = false;
            }
        }
    **/
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
        public void SelectStylesfromDoc(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            if (settings.save_text)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = DocPath();
            openFile.Title = "Select Word Template";
            openFile.FileName = "";
            openFile.Filter = "Word Documents (*.doc;*.docx)|*.doc;*.docx";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (setStyle_frm != null)
                {
                    setStyle_frm.Close();
                    setStyle_frm.Dispose();
                    setStyle_frm = null;
                }
                settings.trackChange(Doc, false);
                setStyle_frm = new setStylesToDoc(Doc, openFile.FileName);
                setStyle_frm.Show();
            }
        }
        public void ChangeStyleSettings(Office.IRibbonControl control)
        {
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (setStyle_frm != null)
            {
                setStyle_frm.Close();
                setStyle_frm.Dispose();
                setStyle_frm = null;
            }
            settings.trackChange(Doc, false);
            setStyle_frm = new setStylesToDoc(Doc, null);
            setStyle_frm.Show();
        }
        public static void CreateEdocFromDoc(Word.Document Doc,String FileName)
        {
            alert = new AlertForm(Doc, 5, 0, FileName);
            alert.Show();
        }
        public static void changeStyles(Word.Document Doc)
        {
            alert = new AlertForm(Doc, 7, 0, "");
            alert.Show();
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
                case "ProcessEdoc_ribbon":
                    {
                        return PictureConverter.ImageToPictureDisp(Properties.Resources.procces_doc);
                    }
                case "CreateEdocFromThis_ribbon":
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
                case "FreeText_ribbon":
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
            }
            return null;

        }
        public bool isActiveAddin()
        {
            if (!Settings.Default.is_active)
            {
                DialogResult dialogResult = MessageBox.Show("License Expired or Out of Date, Enter Additional License?", "License Expired", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    Settings.Default.FirstUse = true;
                    Settings.Default.Save();
                    settings.check_if_vaild_copy();

                }
                return false;
            }
            return true;
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
            if (settings.save_text)
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
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            settings.trackChange(Doc, false);
            alert = new AlertForm(Doc, 3, 0, DocPath() + "//eDoc List of effective Template - A5.docx");
            alert.Show();
        }
        public void OnProcesseDoc(Office.IRibbonControl control)
        {

            if (!isActiveAddin())
                return;
            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (!settings.check_if_edoc(Doc))
            {
                MessageBox.Show("eDoc Only");
                return;
            }

            alert = new AlertForm(Doc, 1 , 1, "");
            alert.Show();
        }
        public void OnHeader1ToTop(Office.IRibbonControl control)
        {

            if (!isActiveAddin())
                return;
            if (settings.save_text)
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
            alert = new AlertForm(Doc, 4, 1, "");
            alert.Show();
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
                    MessageBox.Show("eDoc Only");
                    ribbon.InvalidateControl("toggleButton_ribbon");
                    return;
                }
                settings.save_text = true;
                settings.trackChange(Doc, true);
                alert = new AlertForm(Doc, 6, 1, "");
                alert.Show();
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Do Not Cancel Auto Page Revision - or you will lose all changes.To continue: Press - Yes - and then press Process eDoc ", "Attention !!", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    ribbon.InvalidateControl("toggleButton_ribbon");
                }
                else if (dialogResult == DialogResult.No)
                {
                    settings.save_text = false;
                    ribbon.InvalidateControl("toggleButton_ribbon");
                }
                
            }
            ribbon.InvalidateControl("rev_cbo");
            ribbon.InvalidateControl("date_cbo");

        }
        public bool GetEnable(Office.IRibbonControl control)
        {
            if (settings.save_text)
                return true;
                return false;
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


        public void edit_doc_template(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            if (settings.save_text)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = DocPath();
            openFile.Title = "Select eDocs Word Template";
            openFile.FileName = "";
            openFile.Filter = "Word Documents (*.doc;*.docx)|*.doc;*.docx";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                Globals.ThisAddIn.Application.Documents.Open(openFile.FileName);
        }
        
        public void edit_list_template(Office.IRibbonControl control)
        {

            if (!isActiveAddin())
                return;
            if (settings.save_text)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = DocPath();
            openFile.Title = "Select eDocs List of effective pages Word Template";
            openFile.FileName = "";
            openFile.Filter = "Word Documents (*.doc;*.docx)|*.doc;*.docx";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                Globals.ThisAddIn.Application.Documents.Open(openFile.FileName);
        }
        public void PageRevisionForm(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            if (settings.save_text)
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
                settings.page_rev_frm.rev_date.Text = settings.last_date;

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
        public void FreeText(Office.IRibbonControl control)
        {
            if (!isActiveAddin())
                return;
            if (settings.save_text)
            {
                MessageBox.Show("Auto Rivision on, Please Turn off");
                return;
            }
            if (settings.free_text_frm != null)
            {
                settings.free_text_frm.Close();
                settings.free_text_frm.Dispose();
                settings.free_text_frm = null;
            }

            Doc = Globals.ThisAddIn.Application.ActiveDocument;
            DocSettings DS = new DocSettings(Doc);

            if (settings.check_if_edoc(Doc))
            {
                Doc.Application.ScreenUpdating = false;
                settings.trackChange(Doc, false);
                settings.free_text_frm = new FreeTextfrm(Doc);

                if (settings.last_text1 != "")
                    settings.free_text_frm.text1.Text = settings.last_text1;

                if (settings.last_text2 != "")
                    settings.free_text_frm.text2.Text = settings.last_text2;

                if (settings.last_text3 != "")
                    settings.free_text_frm.text3.Text = settings.last_text3;

                if (settings.last_text4 != "")
                    settings.free_text_frm.text4.Text = settings.last_text4;
                System.Drawing.Point cursor_pos = new System.Drawing.Point(Cursor.Position.X - 50, Cursor.Position.Y - 10);
                settings.free_text_frm.Location = cursor_pos;
                Doc.Application.ScreenUpdating = true;
                settings.trackChange(Doc, true);
                settings.free_text_frm.Show();
            }
            else
            {
                MessageBox.Show("eDoc Only");
            }
        }
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordAddIn2.MyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            if (!Settings.Default.is_active) return;
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
