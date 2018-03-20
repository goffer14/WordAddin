using System;
using System.Collections.Generic;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using BackgroundWorkerDemo;
using eDocs_Editor.Properties;
using MySql.Data.MySqlClient;
using System.Net;
using System.Diagnostics;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Text;
using System.Web;

namespace eDocs_Editor
{
    public static class settings
    {
        public static string doc_password = "edocs_protection";
        public static Dictionary<int, int> original_sections = new Dictionary<int, int> { };
        public static string last_rev = "";
        public static string last_date = "";
        public static string last_text1 = "";
        public static string last_text2 = "";
        public static string last_text3 = "";
        public static string last_text4 = "";
        public static Office.IRibbonControl control_of_list;
        public static PageRevisionFrm page_rev_frm = null;
        public static AuthenticateForm Authenticate_frm = null;
        public static AlertForm alert;
        public static int num_of_click = 0;
        public static bool monitorDoc = false;
        public static DateTime dateToMonitor;
        public static bool View_ShowComments;
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
            try { Doc.Fields.Locked = 0; } catch { return true;};
            return true;
        }
        public static void check_if_vaild_onStartUp()
        {
            if (CheckForInternetConnection())
                logUser();
        }
        public static void process_doc(Word.Document Doc)
        {
            DocSettings DS = new DocSettings(Doc);
            DS.IsAlert = true;
            int DocPageNumber = DS.GetPageNumber(Doc);
            if (DS.PageNumberFromHeaders(DocPageNumber))
                if (settings.monitorDoc)
                {
                    DS.processMonitoring();
                    settings.monitorDoc = false;
                }
            DS.UpDateFields();
            trackChange(Doc, false);
            Doc.Application.ActiveWindow.VerticalPercentScrolled = 0;
            Globals.ThisAddIn.m_Ribbon.ribbon.InvalidateControl("toggleButton_ribbon");
            Globals.ThisAddIn.m_Ribbon.ribbon.InvalidateControl("rev_cbo");
            Globals.ThisAddIn.m_Ribbon.ribbon.InvalidateControl("date_cbo");
        }
        public static void makeAllSameAsPrevious(Word.Document Doc)
        {
            int i;
            if (Doc.Sections.Count > 2)
                for (i = 0; i <= Doc.Sections.Count; i++)
                {
                    try{Doc.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = true;} catch{}
                    try { Doc.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true; } catch { }
                    try { Doc.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true; } catch { }
                    try { Doc.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = true; } catch { }
                    try { Doc.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true; } catch { }
                    try { Doc.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true; } catch { }
                }
            Doc.Application.ActiveWindow.VerticalPercentScrolled = 0;
        }
        public static void logUser()
        {
            try
            {
                var request = (HttpWebRequest)HttpWebRequest.Create("https://global-edocs-auth.herokuapp.com/users/login" + "?id=" + Settings.Default.userId);
                request.Method = "GET";
                request.ContentType = "text/xml; encoding='utf-8'";
                var response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    var rawJson = new StreamReader(response.GetResponseStream()).ReadToEnd();
                    var json = JObject.Parse(rawJson);  //Turns your raw string into a key value lookup
                    bool isActive = json["active"].ToObject<bool>();
                    System.Diagnostics.Debug.WriteLine("isActive - " + isActive);
                    if(isActive)
                        Settings.Default.is_active_auth = "true";
                    else
                        Settings.Default.is_active_auth = "false";
                }
                Settings.Default.Save();
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ex - " + ex);
            }
        }
        public static string getUserData()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("{0}={1}&", "id", HttpUtility.UrlEncode(Settings.Default.userId));
            sb.Remove(sb.Length - 1, 1);
            return sb.ToString();
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
        public static void trackChange(Word.Document Doc, bool boolSum)
        {
            Doc.TrackFormatting = false;
            Doc.TrackMoves = false;
            Doc.TrackRevisions = false;
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
        public static void ChangesExport(Word.Document Doc)
        {
            DocSettings DS = new DocSettings(Doc);
            DS.ChangesExport();
            trackChange(Doc, true);
        }
        public static void CreateMultiLOEP(Word.Document Doc,List<loepDocument> loepDocumentArray)
        {
            DocSettings DS = new DocSettings(Doc);
            DS.CreateMultiLOEP(loepDocumentArray);
            trackChange(Doc, true);
        }
        
        public static void ProcessMonitoring(Word.Document Doc)
        {
            DocSettings DS = new DocSettings(Doc);
            for(int i=1;i<= DS.GetPageNumber(Doc);i++)
                DS.changePageData(i);
            DS.UpDateFields();
        }
        private static int GetPageNumberOfRange(Word.Range range)
        {
            return (int)range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
        }
        public static string LOEPPath(int pageSize)
        {
            string DocPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)+ "\\Global eDocs\\eDocs Add-in\\Templates\\";
            if (pageSize==5)
                DocPath = DocPath +"LOEP5.docx";
            else
                DocPath = DocPath + "LOEPElse.docx";
            return DocPath;
        
        }
        public static void initTOC(Word.Document Doc)
        {
            DocSettings DS = new DocSettings(Doc);
            if (Doc.TablesOfContents.Count < 0)
            {
                MessageBox.Show("Cant Find TOC in DOC");
                return;
            }
            DS.initTOC();
            trackChange(Doc, true);

        }
        public static void init_ListOfE_New(Word.Document RealDoc)
        {
            // JUST FOR ADD PAGES TO
           // BuildTable_ListOfEffctive_BigData(RealDoc, RealDoc.Tables[2]);
            //return;
            //END OF BIG DATA
            Object SelectionNext;
            DocSettings DS = new DocSettings(RealDoc);
            SelectionNext = RealDoc.Application.Selection.End;
            Word.Range rangeForToCopy = RealDoc.Range(ref SelectionNext, ref SelectionNext);
            int pageSize = 4;
            if (rangeForToCopy.PageSetup.PaperSize == Word.WdPaperSize.wdPaperA5)
                pageSize = 5;
            int currentPageNum = Convert.ToInt32(RealDoc.Application.Selection.get_Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber));
            object currentPageNumToRef = currentPageNum;
            Word.Document DocOfE = null;
            try
            {
                DocOfE = Globals.ThisAddIn.Application.Documents.Add(LOEPPath(pageSize));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something Went Wrong - " + ex.Message);
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                openFile.Title = "Select Word Template";
                openFile.FileName = "";
                openFile.Filter = "Word Documents (*.doc;*.docx)|*.doc;*.docx";
                if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    DocOfE = Globals.ThisAddIn.Application.Documents.Add(openFile.FileName);
                }
                else
                    settings.alert.Close();
            }
            Cursor.Current = Cursors.WaitCursor;
            DocOfE.Tables[1].Range.Copy();
            rangeForToCopy.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);
            try
            {
                int DocPageNumber = DS.GetPageNumber(RealDoc);
                RealDoc.Activate();
                int FirstHeaderPage = 0;
                int pageToDo = BuildTable_ListOfEffctivePage(RealDoc, rangeForToCopy.Tables[1], DocPageNumber, FirstHeaderPage);
                copyCellsToTable(rangeForToCopy.Tables[1], DocOfE.Tables[2], (DocPageNumber + pageToDo - FirstHeaderPage), FirstHeaderPage);
                DocOfE.Close(ref saveOption, ref originalFormat, ref routeDocument);
                if (pageToDo > 0)
                {
                    try
                    {
                        string PageString = RealDoc.Variables["edocs_Page" + currentPageNum + "_page"].Value;
                        string PageRev = RealDoc.Variables["edocs_Page" + PageString + "_rev"].Value;
                        string PageDate = RealDoc.Variables["edocs_Page" + PageString + "_date"].Value;
                        for (int i = currentPageNum + 1; i <= currentPageNum + pageToDo; i++)
                        {
                            string PageString2 = RealDoc.Variables["edocs_Page" + i + "_page"].Value;
                            RealDoc.Variables["edocs_Page" + PageString2 + "_rev"].Value = PageRev;
                            RealDoc.Variables["edocs_Page" + PageString2 + "_date"].Value = PageDate;
                        }
                    }
                    catch { }
                }
                if (pageToDo == -1 || !DS.PageNumberFromHeaders(DocPageNumber + pageToDo))
                    RealDoc.Application.ScreenUpdating = true;
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
            Cursor.Current = Cursors.Default;
            rangeForToCopy.Tables[1].Range.Fields.Locked = 0;
            rangeForToCopy.Tables[1].Range.Fields.Update();
            DS.UpDateFields();
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
            RangeToPaste.PasteSpecial(Word.WdPasteDataType.wdPasteRTF);
            RangeToPaste.Paragraphs.SpaceAfter = 0;
            RangeToPaste.Paragraphs.SpaceBefore = 0;
            RangeToPaste.Paragraphs.LeftIndent = 0;
            RangeToPaste.Paragraphs.RightIndent = 0;
            RangeToPaste.Paragraphs.KeepWithNext = 0;
            RangeToPaste.Paragraphs.KeepTogether = 0;
            RangeToCopy.SetRange(ListOfEff_tbl.Rows[leftTbl + 1+ FirstHeaderPage].Range.Start, ListOfEff_tbl.Rows[right].Range.End);
            RangeToCopy.Copy();

            RangeToPaste.SetRange(Realdoc_tbl.Cell(3, 5).Range.Start, Realdoc_tbl.Cell(Realdoc_tbl.Rows.Count-z, 8).Range.End);
            RangeToPaste.Select();
            RangeToPaste.PasteSpecial(Word.WdPasteDataType.wdPasteRTF);
            RangeToPaste.Paragraphs.SpaceAfter = 0;
            RangeToPaste.Paragraphs.SpaceBefore = 0;
            RangeToPaste.Paragraphs.LeftIndent = 0;
            RangeToPaste.Paragraphs.KeepWithNext = 0;
            RangeToPaste.Paragraphs.KeepTogether = 0;
            RangeToPaste.Paragraphs.RightIndent = 0;
            Realdoc_tbl.Cell(1, 1).Range.ListFormat.RemoveNumbers();
        }
        public static int BuildTable_ListOfEffctivePage(Word.Document RealDoc, Word.Table tbl, int PageNum, int FirstHeaderPage)
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
        public static void BuildTable_ListOfEffctive_BigData(Word.Document LOEP, Word.Table tbl)
        {
            Word.Range rng;
            string PageString;
            for (int i=1;i<= tbl.Rows.Count; i++)
            {
                try {
                    PageString = getDataText(i, "date");
                    rng = tbl.Cell(i, 2).Range;
                    rng.Text = "";
                    rng.SetRange(tbl.Cell(i, 2).Range.Start, tbl.Cell(i, 2).Range.Start);
                    LOEP.Fields.Add(rng, Word.WdFieldType.wdFieldDocVariable, PageString, true);

                    PageString = getDataText(i, "rev");
                    rng = tbl.Cell(i, 3).Range;
                    rng.Text = "";
                    rng.SetRange(tbl.Cell(i, 3).Range.Start, tbl.Cell(i, 3).Range.Start);
                    LOEP.Fields.Add(rng, Word.WdFieldType.wdFieldDocVariable, PageString, true);

                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        public static string getDataText(int page,string data)
        {
            return "\"edocs_Page" + page + "_" + data + "\"";
        }

    }
}
