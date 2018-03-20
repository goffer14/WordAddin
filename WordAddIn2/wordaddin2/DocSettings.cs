using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using eDocs_Editor.Properties;
using System.Xml;
using System.Text;
using System.Security.Cryptography;

namespace eDocs_Editor
{

    public class DocSettings
    {
        public Word.Document Doc;
        public object missing = System.Reflection.Missing.Value;
        public object Start;
        public string FooterText;
        public object End;
        public object CurrentPageNumber;
        public object NextPageNumber;
        public object What = Word.WdGoToItem.wdGoToPage;
        public object Which = Word.WdGoToDirection.wdGoToAbsolute;
        public object Miss = System.Reflection.Missing.Value;
        public bool IsAlert;
        public List<object> headingStyles;
        public DocSettings(Word.Document ParentDoc)
        {
            Doc = ParentDoc;
            headingStyles = new List<object>();
            setStylesToDoc();
        }
        public DocSettings()
        {
        }
        public int GetPageNumber(Word.Document doc)
        {
            return doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, System.Reflection.Missing.Value);
        }

        public void InesrtRevDatatoAllHeadingCells(string data, Word.Range range)
        {
            if (data == "page")
                BuildFields(range, "\"edocs_Page_page\"", "\"edocs_Page_" + data + "\"", data);
            else
                BuildFields(range, "\"edocs_Page_page\"", "\"edocs_Page_" + data + "\"", data);
        }
        public void BuildFields(Word.Range range, object PageString, object textString,string data)
        {
            int rangeStart = range.Start;
            range.Text = " ";
            range.End = range.Start;
            range.Fields.Add(range, Word.WdFieldType.wdFieldDocVariable, textString, false);
            if (data != "page")
            {
                range.SetRange(rangeStart + 25, rangeStart + 25);
                range.Fields.Add(range, Word.WdFieldType.wdFieldDocVariable, PageString, true);
            }
            range.Start = range.Start + 25;
            range.Fields.Add(range, Word.WdFieldType.wdFieldEmpty, @"PAGE  \* ARABIC", false);
        }
        public Word.Range GetPageRange(int PageNubmer)
        {
            CurrentPageNumber = (Convert.ToInt32(PageNubmer.ToString()));
            NextPageNumber = (Convert.ToInt32((PageNubmer + 1).ToString()));
            // Get start position of current page
            Start = Doc.GoTo(ref What, ref Which, ref CurrentPageNumber, ref Miss).Start;   
            End = Doc.GoTo(ref What, ref Which, ref NextPageNumber, ref Miss).End;
            // Get text
            if (Convert.ToInt32(Start.ToString()) != Convert.ToInt32(End.ToString()))
                return Doc.Range(ref Start, ref End);
            else
                return Doc.Range(ref Start);
        }
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
        public void insertRev_Rdate(string PageString, string rev, string r_date, string issue, string effective, string text1, string text2, string text3, string text4)
        {
            string page_text = "edocs_Page" + PageString;
            saveDocVariables(page_text, "rev", rev);
            saveDocVariables(page_text, "date", r_date);
            saveDocVariables(page_text, "issue", issue);
            saveDocVariables(page_text, "effective", effective);
            saveDocVariables(page_text, "text1", text1);
            saveDocVariables(page_text, "text2", text2);
            saveDocVariables(page_text, "text3", text3);
            saveDocVariables(page_text, "text4", text4);
        }
        public void saveDocVariables(string page_text, string VariablesText, string value)
        {
            if (value == null)
                return;
            try
            {
                Doc.Variables[page_text + "_"+ VariablesText].Delete();
            }
            catch { }
            Doc.Variables.Add(page_text + "_"+ VariablesText, value);
        }
        public bool PageNumberFromHeaders(int DocPageNumber)
        {
            Word.Range rng = Doc.Range();
            List<header> PageHeader = new List<header>();
            int rngPageNumber = 0;
            for (int i = 1; i <= Doc.Sections.Count; i++)
            {
                try {
                    Doc.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = false;
                    Doc.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].PageNumbers.RestartNumberingAtSection = false;
                    Doc.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers.RestartNumberingAtSection = false;
                }
                catch { }
                try { Doc.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = false; }
                catch { }
            }
            try
            {
                PageHeader = getHeadingArray(headingStyles);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something went wrong: " + rngPageNumber + Environment.NewLine + " Error Msg: " + ex.Message.ToString());
                if (IsAlert)
                    settings.alert.Close();
                return false;
            }
            return InsertPageNumberToDataBase(PageHeader, DocPageNumber);
        }
        public bool InsertPageNumberToDataBase(List<header> PageHeader, int DocPageNumber)
        {
            string valtext;
            string edoc_page_text;
            //string edoc_page_hash_text;
            int nextPageNum;
            int last_page_num = 0;
            int addToINTRO = 1;
            int j;
            int realPage = 0;
            if (PageHeader.Count>0)
                deleteDataBeforeFirstHeader(PageHeader[0].pageNum);
            try
            {

                for (int i = 0; i < PageHeader.Count; i++)
                {
                    if (i == PageHeader.Count - 1)
                    {
                        nextPageNum = DocPageNumber;
                        last_page_num = -1;
                    }
                    else
                        nextPageNum = PageHeader[i + 1].pageNum;
                    for (j = 1 + last_page_num; j <= nextPageNum - PageHeader[i].pageNum; j++)
                    {
                        if (IsAlert && settings.alert.worker.CancellationPending)
                            return false;
                        realPage = ((PageHeader[i].pageNum) + (j - 1) - last_page_num);
                        edoc_page_text = "edocs_Page" + realPage + "_page";
                        if (PageHeader[i].getHeadingString() == "INTRO")
                        {
                            valtext = PageHeader[i].getHeadingString() + " - P-" + addToINTRO;
                            addToINTRO++;
                        }
                        else
                            valtext = PageHeader[i].getHeadingString() + " - P-" + (j - last_page_num).ToString();
                        try{Doc.Variables[edoc_page_text].Delete();}catch{ }
                        Doc.Variables.Add(edoc_page_text, valtext);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something went wrong" + Environment.NewLine + " Error Msg: " + ex.Message.ToString());
                if (IsAlert)
                    settings.alert.Close();
                return false;
            }
            return true;
        }
        public void processMonitoring()
        {
            int numOfChanges = 0;
            string pageString;
            if (Doc.Revisions.Count == 0)
            {
                MessageBox.Show("No tracked changes in this eDoc");
                return;
            }
            System.Diagnostics.Debug.WriteLine("New Pages Rivision: "+ MyRibbon.AutoRevString + " Date: " + MyRibbon.AutoDateString);
            System.Diagnostics.Debug.WriteLine("Num of Changes - " + Doc.Revisions.Count);
            foreach (Word.Revision oRevision in Doc.Revisions)
            {
                if (IsAlert && settings.alert.worker.CancellationPending)
                    return;
                int pageNumberEnd = oRevision.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
                System.Diagnostics.Debug.WriteLine("Change Made in Date  : " + oRevision.Date + " Page - " + pageNumberEnd);
                if (oRevision.Date >= settings.dateToMonitor)
                {
                    Word.Range startRng = Doc.Range(oRevision.Range.Start, oRevision.Range.Start);
                    int pageNumberStrat = startRng.get_Information(Word.WdInformation.wdActiveEndPageNumber);
                    for (int i = 0; i <= pageNumberEnd - pageNumberStrat; i++)
                    {
                        numOfChanges++;
                        pageString = Doc.Variables["edocs_Page" + (pageNumberStrat+i) + "_page"].Value;
                        saveDocVariables("edocs_Page" + pageString, "rev", MyRibbon.AutoRevString);
                        saveDocVariables("edocs_Page" + pageString, "date", MyRibbon.AutoDateString);
                        System.Diagnostics.Debug.WriteLine("Changes Made in page : " + (pageNumberStrat + i));
                    }
                }
                
            }
            if (numOfChanges == 0)
            {
                System.Diagnostics.Debug.WriteLine("No Changes Made from: " + settings.dateToMonitor);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Total num of Changes Made from: " + settings.dateToMonitor + " Are: " + numOfChanges);
            }

        }
        private void deleteDataBeforeFirstHeader(int firstHeadingPage)
        {
            for (int i = 1; i < firstHeadingPage;i++)
                try { Doc.Variables["edocs_Page" + i + "_page"].Delete(); } catch { }
        }
        public List<header> getHeadingArray(List<object> headingStyles)
        {

            List<header> PageHeader = new List<header>();
            int rngPageNumber = 0;
            int firstHeaderPage = 0;
            header midHeader = new header(0, "");
            for (int i=0;i<headingStyles.Count;i++)
            {
                if (headingStyles[i].Equals("Empty"))
                    continue;
                Word.Range rng = Doc.Range();
                rng.Find.set_Style(headingStyles[i]);
                midHeader.pageNum = 0;
                firstHeaderPage = 0;
                while (rng.Find.Execute())
                {
                    if (IsAlert && settings.alert.worker.CancellationPending)
                        return null;
                    rngPageNumber = GetPageNumberOfRange(rng);
                    if (rng.Text.Trim() != "" && rng.Text.Length > 3)
                    {
                        if (rngPageNumber>firstHeaderPage)
                        {
                            if (midHeader.pageNum != 0 && midHeader.pageNum < rngPageNumber)
                            {
                                PageHeader.Add(new header(midHeader.pageNum, midHeader.headingNum));
                                midHeader.pageNum = 0;
                            }
                            else
                                midHeader.pageNum = 0;
                            firstHeaderPage = rngPageNumber;
                            PageHeader.Add(new header(firstHeaderPage, getStringFromHeader(rng)));
                            Doc.Application.Selection.GoTo(ref What, ref Which, firstHeaderPage, ref missing);
                        }
                        else
                        {
                            midHeader.headingNum = getStringFromHeader(rng);
                            midHeader.pageNum = rngPageNumber + 1;
                        }
                        rng.Start = rng.End;
                    }
                    else
                        rng.Start = rng.End;
                }
                if(midHeader.pageNum != 0)
                    PageHeader.Add(new header(midHeader.pageNum, midHeader.headingNum));
                }
            return PageHeader;
        }
        public string getStringFromHeader(Word.Range rng)
        {
            try
            {
                return rng.ListParagraphs[1].Range.ListFormat.ListString.ToString();
            }
            catch
            {
                return "INTRO";
            }
        }
        public void changePageData(int pageNumber)
        {
            int original_section_text = 0;
            string pageString = Doc.Variables["edocs_Page" + pageNumber + "_page"].Value;
            Word.Range rng = GetPageRange(pageNumber);
            settings.original_sections.TryGetValue(pageNumber, out original_section_text);
            if (original_section_text != rng.Text.GetHashCode())
            {
                try
                {
                    Doc.Variables["edocs_Page" + pageString + "_rev"].Delete();
                }
                catch
                { }
                try
                {
                    Doc.Variables["edocs_Page" + pageString + "_date"].Delete();
                }
                catch
                { }
                Doc.Variables.Add("edocs_Page" + pageString + "_rev", MyRibbon.AutoRevString);
                Doc.Variables.Add("edocs_Page" + pageString + "_date", MyRibbon.AutoDateString);
                settings.original_sections.Remove(pageNumber);
                settings.original_sections.Add(pageNumber, rng.Text.GetHashCode());
                System.Diagnostics.Debug.WriteLine("Page: " + pageNumber + " Changed");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Page: " + pageNumber + " Same");
            }
        }
        public void setStylesToDoc()
        {
            addStyle("introduction2");
            addStyle("heading2_name");
            addStyle("appendix2");
        }
        public void addStyle(string styleName)
        {
            try{headingStyles.Add(Doc.Variables[styleName].Value);}
            catch{System.Diagnostics.Debug.WriteLine("Can't Find Style - " + styleName);}
        }
        public void UpDateFields()
        {
            Globals.ThisAddIn.Application.Selection.Fields.Update();
            Globals.ThisAddIn.Application.ActiveDocument.Fields.Update();
            Globals.ThisAddIn.Application.ScreenRefresh();
        }
        public class CustomCell
        {
            public int cellRow;
            public int cellColumn;
            public int cellRowSpan;
            public int cellColumnSpan;
            public String cellText;
        }
        private static int GetPageNumberOfRange(Word.Range range)
        {
            return (int)range.get_Information(Word.WdInformation.wdActiveEndAdjustedPageNumber);
        }
        private static int GetSectionNumberOfRange(Word.Range range)
        {
            return (int)range.get_Information(Word.WdInformation.wdActiveEndSectionNumber);
        }
        public void ChangesExport()
        {
            Word.Document oNewDoc;
            Word.Table oTable;
            Word.Row oRow;
            String strText;
            int n = 0;
            if (Doc.Revisions.Count == 0)
            {
                MessageBox.Show("No tracked changes in this eDoc");
                return;
            }
            oNewDoc = Globals.ThisAddIn.Application.Documents.Add();
            oNewDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            oNewDoc.Content.Text = "";
            oNewDoc.PageSetup.LeftMargin = oNewDoc.Application.CentimetersToPoints(2);
            oNewDoc.PageSetup.RightMargin = oNewDoc.Application.CentimetersToPoints(2);
            oNewDoc.PageSetup.TopMargin = oNewDoc.Application.CentimetersToPoints(3);

            oTable = oNewDoc.Tables.Add(oNewDoc.Application.Selection.Range, 1, 7);
            oNewDoc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "Tracked changes extracted from: " + Doc.FullName + "\r" + "Creation date: " + String.Format("{0:d/M/yyyy HH:mm:ss}", DateTime.Now);

            oTable.AllowAutoFit = false;
            oTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            oTable.PreferredWidth = 100;
            try { oTable.set_Style(Doc.Styles["Table Grid"].NameLocal); }
            catch { }


            foreach (Word.Column oCol in oTable.Columns)
            oCol.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            oTable.Columns[1].PreferredWidth = 5;
            oTable.Columns[2].PreferredWidth = 5;
            oTable.Columns[3].PreferredWidth = 10;
            oTable.Columns[4].PreferredWidth = 35;
            oTable.Columns[5].PreferredWidth = 15;
            oTable.Columns[6].PreferredWidth = 10;
            oTable.Columns[7].PreferredWidth = 10;

            oTable.Rows[1].Cells[1].Range.Text = "Page";
            oTable.Rows[1].Cells[2].Range.Text = "Line";
            oTable.Rows[1].Cells[3].Range.Text = "Type";
            oTable.Rows[1].Cells[4].Range.Text = "What has been inserted or deleted";
            oTable.Rows[1].Cells[5].Range.Text = "Author";
            oTable.Rows[1].Cells[6].Range.Text = "Date";
            oTable.Rows[1].Cells[7].Range.Text = "Time";

            foreach (Word.Revision oRevision in Doc.Revisions)
            {
               
                    if (IsAlert && settings.alert.worker.CancellationPending)
                        return;
                    if (oRevision.Type == Word.WdRevisionType.wdRevisionInsert || oRevision.Type == Word.WdRevisionType.wdRevisionDelete)
                    {
                    try {strText = oRevision.Range.Text; }
                    catch (Exception ex) { strText = "Picture";}
                        n++;
                        oRow = oTable.Rows.Add();
                        oRow.Select();
                        Doc.Application.Selection.GoTo();
                        int pageNumber = oRevision.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
                        int lineNumber = oRevision.Range.get_Information(Word.WdInformation.wdFirstCharacterLineNumber);

                        oRow.Cells[1].Range.Text = pageNumber.ToString();
                        oRow.Cells[2].Range.Text = lineNumber.ToString();
                        if (oRevision.Type == Word.WdRevisionType.wdRevisionInsert)
                        {
                            oRow.Cells[3].Range.Text = "Inserted";
                            oRow.Cells[3].Range.Font.Color = Word.WdColor.wdColorAutomatic;
                        }
                        else
                        {
                            oRow.Cells[3].Range.Text = "Deleted";
                            oRow.Cells[3].Range.Font.Color = Word.WdColor.wdColorRed;
                        }
                        oRow.Cells[4].Range.Text = strText;
                        oRow.Cells[5].Range.Text = oRevision.Author;
                        oRow.Cells[6].Range.Text = oRevision.Date.ToShortDateString();
                        oRow.Cells[7].Range.Text = oRevision.Date.ToShortTimeString();
                    }
                }
            if (n == 0)
            {
                MessageBox.Show("No insertions or deletions were found");
                oNewDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                return;
            }
            oTable.Rows[1].Range.Font.Bold = Convert.ToInt32(true);
            oNewDoc.Application.Selection.GoTo(ref What, ref Which, 1, ref missing);

        }
        public void CreateMultiLOEP(List<loepDocument> loepDocumentArray)
        {
            Word.Table oTable = Doc.Tables.Add(Doc.Application.Selection.Range, 3, 4);
            for (int i = 0; i < loepDocumentArray.Count; i++)
            {
                addDocument(oTable, loepDocumentArray[i]);
                if (i + 1 < loepDocumentArray.Count)
                {
                    oTable.Rows.Add();
                    oTable.Rows.Add();
                    oTable.Rows.Add();
                }
            }

        }
        private void addDocument(Word.Table oTable, loepDocument document)
        {
            try { oTable.set_Style(Doc.Styles["Table Grid"].NameLocal); }
            catch { }
            oTable.Range.ParagraphFormat.SpaceAfter = 0;

            int fileExtPos = document.name.LastIndexOf(".");
            if (fileExtPos >= 0)
                document.name = document.name.Substring(0, fileExtPos);
            oTable.Rows[oTable.Rows.Count-2].Range.Text = document.name;
            oTable.Rows[oTable.Rows.Count - 2].Range.Font.Bold = 1;
            oTable.Rows[oTable.Rows.Count - 2].Range.Font.Size = 10;
            oTable.Rows[oTable.Rows.Count - 2].Range.Font.Position = 1;
            oTable.Rows[oTable.Rows.Count - 2].Cells[1].Merge(oTable.Rows[oTable.Rows.Count - 2].Cells[2]);
            oTable.Rows[oTable.Rows.Count - 2].Cells[1].Merge(oTable.Rows[oTable.Rows.Count - 2].Cells[2]);
            oTable.Rows[oTable.Rows.Count - 2].Cells[1].Merge(oTable.Rows[oTable.Rows.Count - 2].Cells[2]);
            oTable.Cell(oTable.Rows.Count - 2, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

            oTable.Rows[oTable.Rows.Count- 1].Range.Font.Size = 10;
            oTable.Rows[oTable.Rows.Count].Range.Font.Size = 9;
            oTable.Rows[oTable.Rows.Count - 1].Range.Font.Bold = 1;
            oTable.Cell(oTable.Rows.Count - 1, 1).Range.Text = "From Page";
            oTable.Cell(oTable.Rows.Count - 1, 2).Range.Text = "To Page";
            oTable.Cell(oTable.Rows.Count - 1, 3).Range.Text = "Revision";
            oTable.Cell(oTable.Rows.Count - 1, 4).Range.Text = "Date";
            oTable.Cell(oTable.Rows.Count - 1, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oTable.Cell(oTable.Rows.Count - 1, 2).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oTable.Cell(oTable.Rows.Count - 1, 3).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oTable.Cell(oTable.Rows.Count - 1, 4).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

            oTable.Cell(oTable.Rows.Count, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oTable.Cell(oTable.Rows.Count, 2).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oTable.Cell(oTable.Rows.Count, 3).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oTable.Cell(oTable.Rows.Count, 4).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

            addPages(document.location, oTable);
        }
        private void addPages(string docLoc, Word.Table oTable)
        {
            Microsoft.Office.Interop.Word.Application word = null;
            word = new Microsoft.Office.Interop.Word.Application();

            object inputFile = docLoc;
            object confirmConversions = false;
            object readOnly = true;
            object visible = false;
            object missing = Type.Missing;

            // Open the document...
            Microsoft.Office.Interop.Word.Document tempDoc = null;
            tempDoc = word.Documents.Open(
                ref inputFile, ref confirmConversions, ref readOnly, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref visible,
                ref missing, ref missing, ref missing, ref missing);
            tempDoc.Activate();

            string tempPageString = "";
            string tempPageRevisionText = "";
            string tempPageDateText = "";
            List<multiLOEP> pageMultiLOEP = new List<multiLOEP>();
            for (int i = 1; i <= GetPageNumber(tempDoc); i++)
            {
                tempPageString = getString(tempDoc, "edocs_Page" + i + "_page");
                tempPageRevisionText = getString(tempDoc, "edocs_Page" + tempPageString + "_rev");
                tempPageDateText = getString(tempDoc, "edocs_Page" + tempPageString + "_date");
                if (pageMultiLOEP.Count == 0)
                    pageMultiLOEP.Add(new multiLOEP(tempPageString, tempPageRevisionText, tempPageDateText));
                else if(pageMultiLOEP[pageMultiLOEP.Count - 1].rev != tempPageRevisionText)
                {
                    if(addRows(oTable.Rows[oTable.Rows.Count], pageMultiLOEP[pageMultiLOEP.Count-1]))
                        oTable.Rows.Add();
                    pageMultiLOEP.Add(new multiLOEP(tempPageString, tempPageRevisionText, tempPageDateText));
                }
                else
                    pageMultiLOEP[pageMultiLOEP.Count-1].endPage = tempPageString;
            }
            addRows(oTable.Rows[oTable.Rows.Count], pageMultiLOEP[pageMultiLOEP.Count-1]);
            tempDoc.Close(null, null, null);
            word.Quit(null, null, null);
            word = null;
            GC.Collect();
            System.Diagnostics.Debug.WriteLine("Finish Multi");
        }
        private string getString(Word.Document tempDoc,string str)
        {
            string tempString= "Empty";
            try { tempString = tempDoc.Variables[str].Value; } catch {return tempString;}
            return tempString;
        }
        private bool addRows(Word.Row oRow, multiLOEP tempRow)
        {
            if (tempRow.startPage == "Empty")
                return false;
            oRow.Cells[1].Range.Text = tempRow.startPage;
            oRow.Cells[2].Range.Text = tempRow.endPage;
            oRow.Cells[3].Range.Text = tempRow.rev;
            oRow.Cells[4].Range.Text = tempRow.date;
            return true;
        }
        public void initTOC()
        {
            System.Diagnostics.Debug.WriteLine("Paragraph Count - " + Doc.TablesOfContents[1].Range.Paragraphs.Count);
            //foreach (Word.Paragraph pra in Doc.TablesOfContents[1].Range.Paragraphs)
            for (int i=2;i<= Doc.TablesOfContents[1].Range.Paragraphs.Count;i++)
                        { 
                    Word.Paragraph pra = Doc.TablesOfContents[1].Range.Paragraphs[i];
                    if (IsAlert && settings.alert.worker.CancellationPending)
                        return;
                    string rangeText = pra.Range.Text;
                    string[] numbers = Regex.Split(rangeText, @"\D+");
                    System.Diagnostics.Debug.WriteLine("Paragraph " + rangeText);

                    for (int z = numbers.Length - 1; z >= 0; z--)
                    {
                        if (!string.IsNullOrEmpty(numbers[z]))
                        {
                            int q = int.Parse(numbers[z]);
                            replaceTextInTOc(pra, q);
                            //replaceTextInTOc2(pra, q);
                            break;
                        }
                    }
            }
        }
        public void replaceTextInTOc2(Word.Paragraph par, int findText)
        {
            try
            {
                string final;
                string temp = par.Range.Text;           
                string pageValue = null;
                //Word.ParagraphFormat parFormat = par.Format;
                try
                {
                    pageValue = Doc.Variables["edocs_Page" + findText + "_page"].Value;
                    //pageValue = Regex.Replace(pageValue, @"\s", "");
                }
                catch (Exception e)
                { 
                    System.Diagnostics.Debug.WriteLine(e.Message);
                    return;
                }

                if (temp.Length > 1)
                {
                    final = ReplaceAt(temp, temp.Length - 2 - findText.ToString().Length+1, findText.ToString().Length, pageValue);
                    if (final != temp)
                    {
                        par.Range.Text = final;
                        //par.Format = parFormat;
                    }
                }
                
            }
            catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
        }
        public static string ReplaceAt(string str, int index, int length, string replace)
        {
            return str.Remove(index, Math.Min(length, str.Length - index))
                    .Insert(index, replace);
        }
        public void replaceTextInTOc(Word.Paragraph pra,int pageNum)
        {
            object findText = pageNum;
            string pageValue = null;
            try
            {
                pageValue = Doc.Variables["edocs_Page" + pageNum + "_page"].Value;
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
                return;
            }
            pra.Range.Find.ClearFormatting();
            object replaceText = pageValue;
                object replaceAll = Word.WdReplace.wdReplaceOne;
                object forward = false;
                object matchAllWord = false;
            if (pra.Range.Find.Execute(ref findText, ref missing, ref matchAllWord, ref missing, ref missing, ref missing, ref forward, Word.WdFindWrap.wdFindAsk, ref missing, ref replaceText, ref replaceAll, ref missing, ref missing, ref missing, ref missing))
                    System.Diagnostics.Debug.WriteLine("FOUND - " + pageNum);
                else
                    System.Diagnostics.Debug.WriteLine("Not FOUND");
            

        }
    }
}
