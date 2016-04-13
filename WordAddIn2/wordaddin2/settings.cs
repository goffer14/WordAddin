using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Cryptography;

using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Drawing;
using System.ComponentModel;
using System.IO;
using BackgroundWorkerDemo;
using WordAddIn2.Properties;
using MySql.Data.MySqlClient;
using System.Net;

namespace WordAddIn2
{
    public static class settings
    {
        public static string doc_password = "edocs_protection";
        public static Dictionary<string, string> original_sections = new Dictionary<string, string> { };
        public static string last_rev = "";
        public static string last_date = "";
        public static string last_text1 = "";
        public static string last_text2 = "";
        public static string last_text3 = "";
        public static string last_text4 = "";
        public static Office.IRibbonControl control_of_list;
        public static PageRevisionFrm page_rev_frm = null;
        public static AuthenticateForm Authenticate_frm = null;
        public static FreeTextfrm free_text_frm = null;
        public static AlertForm alert;
        public static int num_of_click = 0;
        public static bool View_ShowComments;
        public static bool save_text=false;
        public static bool View_ShowRevisionsAndComments;
        public static bool with_page_zero = true;
        public static bool TrackFormatting;
        public static bool TrackMoves;
        public static bool TrackRevisions;
        public static object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
        public static object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
        public static object what = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage;
        public static object which = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst;
        public static object missing = System.Reflection.Missing.Value;
        public static object routeDocument = false;
        public static bool check_if_edoc(Word.Document Doc)
        {
            int numOfData = 0;
            foreach (Word.Variable varr in Doc.Variables)
                if (varr.Name.Contains("edocs"))
                    numOfData++;

            if (numOfData >= 12)
                return true;
            return false;
        }

        public static void check_if_vaild_copy()
        {
            if (Settings.Default.FirstUse)
            {
                    Authenticate_frm = new AuthenticateForm();
                    Authenticate_frm.Show();
                    return;
            }
            else if(CheckForInternetConnection())
            {
                updateLicense();
            }
            if (Settings.Default.Last_connction.Add(new TimeSpan(10, 0, 0, 0)) < DateTime.Now.Date)
                Settings.Default.is_active = false;
            if (Settings.Default.StartTime.Add(new TimeSpan(Settings.Default.days_for_use, 0, 0, 0)) < DateTime.Now.Date)
                Settings.Default.is_active = false;

            Settings.Default.Save();
        }
        public static void updateLicense()
        {
            string s;
            MySqlConnection mcon = new MySqlConnection(Settings.Default.serverString);
            MySqlCommand mcd;
            MySqlDataReader mdr;
            DateTime CurrentDate = DateTime.Now.Date;
            try
            {
                mcon.Open(); s = "select * from license_table where license_key = '" + Settings.Default.addin_license + "'";
                mcd = new MySqlCommand(s, mcon);
                mdr = mcd.ExecuteReader();
                if (mdr.Read())
                {
                    int is_active = mdr.GetInt32("is_active");
                    mcon.Close();
                    if (is_active == 1)
                        Settings.Default.is_active = true;
                    else
                        Settings.Default.is_active = false;
                }
                else
                {
                    Settings.Default.is_active = false;
                }
                Settings.Default.Last_connction = CurrentDate;
                Settings.Default.Save();
            }
            catch (Exception ex)
            {

            }
            
        }
        public static bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                {
                    using (var stream = client.OpenRead("http://www.google.com"))
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }
        }
        public static string[] GetCellFromPage(Word.Document Doc, int pageNum)
        {

            string[] vars = new string[3];
            string letssee;
            try
            {
                vars[0] = Doc.Variables["edocs_PAGE" + pageNum + "_page"].Value;
            }
            catch
            {
                vars[0] = "No Page";
                vars[1] = "No Revision";
                vars[2] = "No Date";
                return vars;
            }
            try
            {
                letssee = "edocs_PAGE" + vars[0] + "_rev";
                vars[1] = Doc.Variables[letssee].Value;
            }
            catch
            {
                vars[1] = "No Revision";
            }
            try
            {
                vars[2] = Doc.Variables["edocs_PAGE" + vars[0] + "_date"].Value;
            }
            catch
            {
                vars[2] = "No Date";
            }
            return vars;

        }
  
        public static bool init_doc(Word.Document Doc)
        {
            DocSettings DS = new DocSettings(Doc);
            if (!DS.dt_init())
                return false;
            DS.InsertCodeToHeader();
            DS.InesrtDatatoHeadingCells();
            if(DS.insertSectionBreakBeforeHeading1())
              DS.resetHeadersInFirstSection();
            DS.UpDateFields();
            return true;
        }
        public static void ChagneStyles(Word.Document Doc)
        {
            DocSettings DS = new DocSettings(Doc);
            DS.InsertCodeToHeader();
            if (Doc.Sections.Count>1)
                DS.resetHeadersInFirstSection();
            DS.UpDateFields();
        }

        public static void trackChange(Word.Document Doc, bool boolSum)
        {
            Doc.TrackFormatting = false;
            Doc.TrackMoves = false;
            Doc.TrackRevisions = false;

        }
        public static void OnCreateEdocFromDoc(Word.Document CopyTo, string open_file)
        {
            Word.Document TeamlpateToCopy = Globals.ThisAddIn.Application.Documents.Add(open_file);
            object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
            object routeDocument = false;
            if(!replacSectionBreak(CopyTo))
            {
                //MessageBox.Show("In Order to Proceed Delete all Section Breaks in Doc");
                TeamlpateToCopy.Close(ref saveOption, ref originalFormat, ref routeDocument);
                settings.alert.Close();
                return;
            }
            for (int i = 1; i <= CopyTo.Sections.Count; i++)
            {
                TeamlpateToCopy.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Tables[1].Range.Copy();
                CopyTo.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Application.Selection.WholeStory();
                CopyTo.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paste();
                CopyTo.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = false;
                try
                {
                    TeamlpateToCopy.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Tables[1].Range.Copy();
                    CopyTo.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Application.Selection.WholeStory();
                    CopyTo.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paste();
                    CopyTo.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = false;
                }
                catch {}
            }
            TeamlpateToCopy.Close(ref saveOption, ref originalFormat, ref routeDocument);
            CopyTo.Application.Selection.Collapse();
            if (!settings.init_doc(CopyTo))
            {
                MessageBox.Show("Something Went Wrong");
                settings.alert.Close();
                return;
            }

            CopyTo.Application.ActiveWindow.VerticalPercentScrolled = 0;
            CopyTo.Application.Selection.Collapse();
            CopyTo.Application.ActiveWindow.View.ShowFieldCodes = false;
        }
        public static void changeStyles(Word.Document CopyTo)
        {
            settings.ChagneStyles(CopyTo);
            CopyTo.Application.ActiveWindow.VerticalPercentScrolled = 0;
            CopyTo.Application.Selection.Collapse();
            CopyTo.Application.ActiveWindow.View.ShowFieldCodes = false;
        }
        public static bool replacSectionBreak(Word.Document Doc)
        {
            object findText = "^b";
            Word.Range rng = Doc.Range();
            while (rng.Find.Execute(findText))
            {

                Object beginPageNext = rng.Start;
                Object endPageNext = rng.End;
                Word.Range rangeForBreak = Doc.Range(ref beginPageNext, ref endPageNext);
                rangeForBreak.Select();
                rangeForBreak.InsertBreak(Word.WdBreakType.wdPageBreak);
                rng = Doc.Range();
            }
            if (Doc.Sections.Count > 1)
                return false;
            return true;
        }
        public static int GetPageNumber(Word.Document doc)
        {
            return doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, System.Reflection.Missing.Value);
        }
        public static void process_doc(Word.Document Doc)
        {
            DocSettings DS = new DocSettings(Doc);
            DS.IsAlert = true;
            int DocPageNumber = DS.GetPageNumber(Doc);
            if (DS.PageNumberFromHeaders(DocPageNumber))
                DS.UpDateFields();
            trackChange(Doc, true);
            save_text = false;
            Globals.ThisAddIn.m_Ribbon.ribbon.InvalidateControl("toggleButton_ribbon");
            Globals.ThisAddIn.m_Ribbon.ribbon.InvalidateControl("rev_cbo");
            Globals.ThisAddIn.m_Ribbon.ribbon.InvalidateControl("date_cbo");
            Doc.Application.ActiveWindow.VerticalPercentScrolled = 0;

        }
        public static void Header1ToTop(Word.Document Doc)
        {

            // Blind the Application to do the magic behind the curtains!
            //Doc.Application.ScreenUpdating = false;
            DocSettings DS = new DocSettings(Doc);
            Doc.Application.ScreenUpdating = true;
            DS.arrangeHeader1();
            trackChange(Doc, true);
            Doc.Application.ActiveWindow.VerticalPercentScrolled = 0;

        }
        public static void SavePagesText(Word.Document Doc)
        {

            // Blind the Application to do the magic behind the curtains!
            //Doc.Application.ScreenUpdating = false;
            if (!save_text)
                return;

                DocSettings DS = new DocSettings(Doc);
                save_text = true;
                original_sections = DS.saveAllRangsText();
            if (original_sections == null)
            {
                save_text = false;
                Globals.ThisAddIn.m_Ribbon.ribbon.InvalidateControl("toggleButton_ribbon");
                Globals.ThisAddIn.m_Ribbon.ribbon.InvalidateControl("rev_cbo");
                Globals.ThisAddIn.m_Ribbon.ribbon.InvalidateControl("date_cbo");
            }
                Doc.Application.ActiveWindow.VerticalPercentScrolled = 0;
            
        }
        public static void init_ListOfE_New(Word.Document RealDoc, string fileName)
        {
            Object SelectionNext;
            DocSettings DS = new DocSettings(RealDoc);
            SelectionNext = RealDoc.Application.Selection.End;
            Word.Range rangeForToCopy = RealDoc.Range(ref SelectionNext, ref SelectionNext);
            int currentPageNum = Convert.ToInt32(RealDoc.Application.Selection.get_Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber));
            object currentPageNumToRef = currentPageNum;
            Word.Document DocOfE = Globals.ThisAddIn.Application.Documents.Add(fileName);
            
            // JUST FOR ADD PAGES TO
            // BuildTable_ListOfEffctivePage_BigData(DocOfE, DocOfE.Tables[2], 1500);
            //return;
            //END OF BIG DATA
            DocOfE.Tables[1].Range.Copy();
            rangeForToCopy.Paste();
            try
            {
                int DocPageNumber = DS.GetPageNumber(RealDoc);
                RealDoc.Activate();
                int FirstHeaderPage = DS.GetFirstHeader()[1]-1;
                int pageToDo = BuildTable_ListOfEffctivePage(RealDoc, rangeForToCopy.Tables[1], DocPageNumber, FirstHeaderPage);
                copyCellsToTable(rangeForToCopy.Tables[1], DocOfE.Tables[2], (DocPageNumber + pageToDo- FirstHeaderPage), FirstHeaderPage);
                DocOfE.Close(ref saveOption, ref originalFormat, ref routeDocument);

                if (pageToDo > 0)
                {
                    string PageString = RealDoc.Variables["edocs_PAGE" + currentPageNum + "_page"].Value;
                    string PageRev = RealDoc.Variables["edocs_PAGE" + PageString + "_rev"].Value;
                    string PageDate = RealDoc.Variables["edocs_PAGE" + PageString + "_date"].Value;
                    int pageNum = Convert.ToInt32(PageString.Substring(PageString.IndexOf("P-") + 2));
                    PageString = PageString.Substring(0, PageString.IndexOf("P-") + 2);
                    for (int i = 1; i <= pageToDo; i++)
                    {
                        string EndPageString = PageString + (pageNum + i).ToString();
                        RealDoc.Variables["edocs_PAGE" + currentPageNum + i + "_page"].Value = EndPageString;
                        RealDoc.Variables["edocs_PAGE" + EndPageString + "_rev"].Value = PageRev;
                        RealDoc.Variables["edocs_PAGE" + EndPageString + "_date"].Value = PageDate;
                    }
                }
                if (pageToDo == -1 || !DS.PageNumberFromHeaders(DocPageNumber+ pageToDo))
                    return;
                DS.UpDateFields();
                RealDoc.Application.Selection.GoTo(ref what, ref which, ref currentPageNumToRef, ref missing);
                trackChange(RealDoc, true);
            }
            catch (Exception ex)
            {

                trackChange(RealDoc, true);
                RealDoc.Application.ScreenUpdating = true;
                MessageBox.Show("Something Went Wrong - " + ex.Message);
                settings.alert.Close();
            }
        }
        private static int GetPageNumberOfRange(Word.Range range)
        {
            return (int)range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
        }
        public static void copyCellsToTable(Word.Table Realdoc_tbl, Word.Table ListOfEff_tbl, int pageToDo,int FirstHeaderPage)
        {
            int leftTbl = pageToDo / 2;
            int z = 0;
            int right = pageToDo+ FirstHeaderPage;
            if (pageToDo % 2 != 0)
            {
                leftTbl = leftTbl + 1;
                z = 1;
            }
            //left side
            Word.Range RangeToCopy = ListOfEff_tbl.Range;
            RangeToCopy.SetRange(ListOfEff_tbl.Rows[FirstHeaderPage+1].Range.Start, ListOfEff_tbl.Rows[leftTbl+ FirstHeaderPage].Range.End);
            RangeToCopy.Copy();
            Word.Range RangeToPaste = Realdoc_tbl.Range;
            RangeToPaste.SetRange(Realdoc_tbl.Cell(3, 1).Range.Start, Realdoc_tbl.Cell(Realdoc_tbl.Rows.Count, 4).Range.End);
            RangeToPaste.Select();
            RangeToPaste.Paste();
            RangeToCopy.SetRange(ListOfEff_tbl.Rows[leftTbl + 1+ FirstHeaderPage].Range.Start, ListOfEff_tbl.Rows[right].Range.End);
            RangeToCopy.Copy();

            RangeToPaste.SetRange(Realdoc_tbl.Cell(3, 5).Range.Start, Realdoc_tbl.Cell(Realdoc_tbl.Rows.Count-z, 8).Range.End);
            RangeToPaste.Select();
            RangeToPaste.Paste();
        }
        public static int BuildTable_ListOfEffctivePage(Word.Document RealDoc, Word.Table tbl, int PageNum,int FirstHeaderPage)
        {
            DocSettings DS = new DocSettings(RealDoc);
            int pagesSum = 1;
            PageNum = PageNum - FirstHeaderPage;
            int startPageNumber = PageNum;
            int RowToadd=0;
           
                while (pagesSum * 2 < PageNum)
                {
                if (settings.alert.worker.CancellationPending)
                    return -1;
                RowToadd = (PageNum / 2)- pagesSum;
                if (PageNum % 2 != 0)
                    RowToadd = RowToadd+1;

                    tbl.Rows[tbl.Rows.Count].Select();
                    RealDoc.ActiveWindow.Selection.InsertRowsBelow(RowToadd);
                pagesSum = tbl.Rows.Count-2;
                PageNum = DS.GetPageNumber(RealDoc) - FirstHeaderPage;
                }
            
            if (PageNum % 2 != 0)
            {
                tbl.Cell(tbl.Rows.Count, 5).Range.Text = "x";
                tbl.Cell(tbl.Rows.Count, 6).Range.Text = "x";
                tbl.Cell(tbl.Rows.Count, 7).Range.Text = "x";
                tbl.Cell(tbl.Rows.Count, 8).Range.Text = "x";
            }
            return PageNum - startPageNumber;
        }
        public static void insertDocVirableToRange(Word.Document RealDoc, Word.Table tbl, int pageNum)
        {
            //1 - chepter , 2 - page , 3-date , 4-rev
            Word.Range rng;
            object PageString;
            object StringJob;

            tbl.Cell(pageNum, 1).Range.Text = pageNum.ToString();

            PageString = "\"edocs_PAGE" + pageNum + "_page\"";
            rng = tbl.Cell(pageNum, 2).Range;
            rng.SetRange(tbl.Cell(pageNum, 2).Range.Start, tbl.Cell(pageNum, 2).Range.Start);
            RealDoc.Fields.Add(rng, Word.WdFieldType.wdFieldDocVariable, PageString, true);

            tbl.Cell(pageNum, 2).Range.Fields.Update();

            StringJob = "\"edocs_PAGE_date\"";
            rng.SetRange(tbl.Cell(pageNum, 3).Range.Start, tbl.Cell(pageNum, 3).Range.Start);
            RealDoc.Fields.Add(rng, Word.WdFieldType.wdFieldDocVariable, StringJob, true);
            rng.Start = rng.Start + 25;
            RealDoc.Fields.Add(rng, Word.WdFieldType.wdFieldDocVariable, PageString, true);

            tbl.Cell(pageNum, 3).Range.Fields.Update();

            StringJob = "\"edocs_PAGE_rev\"";
            rng.SetRange(tbl.Cell(pageNum, 4).Range.Start, tbl.Cell(pageNum, 4).Range.Start);
            RealDoc.Fields.Add(rng, Word.WdFieldType.wdFieldDocVariable, StringJob, true);
            rng.Start = rng.Start + 25;
            RealDoc.Fields.Add(rng, Word.WdFieldType.wdFieldDocVariable, PageString, true);

            tbl.Cell(pageNum, 4).Range.Fields.Update();


        }
        public static void add_new_page(Word.Document Doc)
        {
            Object beginPageNext;
            beginPageNext = Doc.Application.Selection.Range.End;
            Word.Range rangeForBreak = Doc.Range(ref beginPageNext);
            rangeForBreak.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            Word.HeaderFooter header_footer = rangeForBreak.Sections.First.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            header_footer.LinkToPrevious = false;
        }
        
    }
}
