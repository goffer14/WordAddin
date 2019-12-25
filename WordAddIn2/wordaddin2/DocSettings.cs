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
        public DocSettings(Word.Document ParentDoc,String WTF)
        {
            Doc = ParentDoc;
        }
        public int GetPageNumber(Word.Document doc)
        {
            return doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, System.Reflection.Missing.Value);
        }

        public int GetPageNumber()
        {
            return Doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, System.Reflection.Missing.Value);
        }

        public void InesrtRevDatatoAllHeadingCells(string data, Word.Range range)
        {
                BuildFields(range, "\"edocs_Page_page\"", "\"edocs_Page_" + data + "\"", data);
        }
        public void buildIfEmptyField(Word.Range rngTarget, object PageString, object textString, string data)
        {
            string sQ = '"'.ToString();
            Word.Field fldIf = null;
            Word.View vw = Doc.ActiveWindow.View;
            rngTarget.Text = " ";
            rngTarget.End = rngTarget.Start;
            bool bViewFldCodes = false;
            string sFieldCode;
            string pageVar = "{PAGE  \\* ARABIC}";
            string docVar = "{DOCVARIABLE " + sQ + "edocs_Page" +pageVar+"_page" + sQ + "\\* MERGEFORMAT}";
            if (data == "page")
                docVar = pageVar;
            string fullPageVar = "{DOCVARIABLE " + sQ + "edocs_Page" + docVar + "_" + data+ sQ + "}";
            string ifVar = "IF" + fullPageVar + "=" + sQ + "Error!*" + sQ + " " + sQ + "eDoc Empty Field" + sQ + fullPageVar;

            System.Diagnostics.Debug.WriteLine("String - " + ifVar);

            bViewFldCodes = vw.ShowFieldCodes;
            //Finding text in a field codes requires field codes to be shown
            if (!bViewFldCodes) vw.ShowFieldCodes = true;


            fldIf = rngTarget.Fields.Add(rngTarget, Word.WdFieldType.wdFieldEmpty, ifVar, false);
            sFieldCode = GenerateNestedField(fldIf, fullPageVar,true);
            sFieldCode = GenerateNestedField(fldIf, fullPageVar, false);
            if (data != "page")
            {
                sFieldCode = GenerateNestedField(fldIf, docVar, true);
                sFieldCode = GenerateNestedField(fldIf, docVar, false);
            }
            sFieldCode = GenerateNestedField(fldIf, pageVar,true);
            sFieldCode = GenerateNestedField(fldIf, pageVar, false);
            rngTarget.Fields.Update();
            vw.ShowFieldCodes = bViewFldCodes;

        }
        private string GenerateNestedField(Word.Field fldOuter, string sPlaceholder, object forward)
        {
            Word.Range rngFld = fldOuter.Code;
            bool bFound;
            string sFieldCode;

            //Get the field code from the placeholder by removing the { }
            sFieldCode = sPlaceholder.Substring(1, sPlaceholder.Length - 2); //Mid(sPlaceholder, 2, Len(sPlaceholder) - 2)
            rngFld.TextRetrievalMode.IncludeFieldCodes = true;

            System.Diagnostics.Debug.WriteLine("sFieldCode: " + sFieldCode);
            System.Diagnostics.Debug.WriteLine("rngFld.StoryType.ToString(): " + rngFld.StoryType.ToString());
            System.Diagnostics.Debug.WriteLine("rngFld.StoryType.ToString(): " + rngFld.Text.ToString());
            rngFld.Fields.Update();
            rngFld.Find.ClearFormatting();
            bFound = rngFld.Find.Execute(sPlaceholder, ref missing, ref missing, ref missing, ref missing, ref missing, ref forward, Word.WdFindWrap.wdFindContinue, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            System.Diagnostics.Debug.WriteLine("bFound: " + bFound);
            System.Diagnostics.Debug.WriteLine(" ");
            if (bFound) Doc.Fields.Add(rngFld, Word.WdFieldType.wdFieldEmpty, sFieldCode, false);

            return fldOuter.Code.ToString();
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
        public bool PageNumberFromHeaders(int DocPageNumber,string type)
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
            if (type.Equals("styles"))
            {
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
            }
            else
            {
                    PageHeader.Add(new header(1, ""));
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
                            addPageTemplates(valtext, PageHeader[i].getHeadingString(), addToINTRO, nextPageNum - PageHeader[i].pageNum);
                            addToINTRO++;
                        }
                        else
                        {
                            if (PageHeader[i].getHeadingString().Length > 0)
                                valtext = PageHeader[i].getHeadingString() + " - P-" + (j - last_page_num).ToString();
                            else
                                valtext = "P" + (j - last_page_num).ToString();
                            addPageTemplates(valtext, PageHeader[i].getHeadingString(), (j - last_page_num), nextPageNum - (PageHeader[i].pageNum+last_page_num));
                        }
                        try {Doc.Variables[edoc_page_text].Delete();}catch{ }
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
        private void addPageTemplates(string edoc_page_text,string headingString,int page ,int totalNumnber)
        {
            //X-P1
            string pageTemplateEdocCPx = "edocs_Page" + edoc_page_text + "_X-P1";
            string valtext;
            if (headingString == "INTRO")
                valtext = headingString + "-P" + page;
            else if (headingString.Length > 0)
                valtext = headingString + "-P" + page;
            else
                valtext = "P" + page;
            try { Doc.Variables[pageTemplateEdocCPx].Delete(); } catch { }
            Doc.Variables.Add(pageTemplateEdocCPx, valtext);

            //X-P-1
            string pageTemplateEdocC_P_x = "edocs_Page" + edoc_page_text + "_X-P-1"; 
            if (headingString == "INTRO")
                valtext = headingString + "-P-" + page;
            else if (headingString.Length > 0)
                valtext = headingString + "-P-" + page;
            else
                valtext = "P-" + page;
            try { Doc.Variables[pageTemplateEdocC_P_x].Delete(); } catch { }
            Doc.Variables.Add(pageTemplateEdocC_P_x, valtext);

            //X1
            string pageTemplateEdocX = "edocs_Page" + edoc_page_text + "_X1";
            if (headingString == "INTRO")
                valtext = headingString + "-" + page;
            else if (headingString.Length > 0)
                valtext = headingString + "-" + page;
            else
                valtext = page.ToString();
            try { Doc.Variables[pageTemplateEdocX].Delete(); } catch { }
            Doc.Variables.Add(pageTemplateEdocX, valtext);

            //Total Pages
            string totalNumberOfpages = "edocs_Page" + edoc_page_text + "_total_pages";
            valtext = "of " + totalNumnber;
            try { Doc.Variables[totalNumberOfpages].Delete(); } catch { }

            System.Diagnostics.Debug.WriteLine("Pages - " + valtext);

            Doc.Variables.Add(totalNumberOfpages, valtext);


        }
        private void addEmptyFiled(int pageNumber)
        {
            setVarString("page", pageNumber);
            setVarString("date", pageNumber);
            setVarString("rev", pageNumber);
            setVarString("issue", pageNumber);
            setVarString("effective", pageNumber);
            setVarString("text1", pageNumber);
            setVarString("text2", pageNumber);
            setVarString("text3", pageNumber);
            setVarString("text4", pageNumber);

        }
        public void moveToNewVersion()
        {
            foreach (Word.Variable var in Doc.Variables)
            {
                if (var.Value.Contains(" P-"))
                {
                    System.Diagnostics.Debug.WriteLine("Old String - " + var.Value);
                    String oldPageVar = var.Value;
                    string oldValue = var.Value;
                    string newValue = var.Value.Replace(" P-", " P");
                    updatePageVars(newValue, oldValue);
                    var.Value = newValue;
                    System.Diagnostics.Debug.WriteLine("New String - " + var.Value);
                }

            }
           UpDateFields();
        }
        private void updatePageVars(string newPageText,string OldPageText)
        {
            setNewValue("date", newPageText, OldPageText);
            setNewValue("rev", newPageText, OldPageText);
            setNewValue("issue", newPageText, OldPageText);
            setNewValue("effective", newPageText, OldPageText);
            setNewValue("text1", newPageText, OldPageText);
            setNewValue("text2", newPageText, OldPageText);
            setNewValue("text3", newPageText, OldPageText);
            setNewValue("text4", newPageText, OldPageText);
        }
        private void setNewValue(string VariablesText, string newPageText, string OldPageText)
        {
            string oldValue;
            try
            {
                System.Diagnostics.Debug.WriteLine("OldPageText + VariablesText" + OldPageText + "_" + VariablesText);
                oldValue = Doc.Variables["edocs_Page"+OldPageText + "_" + VariablesText].Value;
                System.Diagnostics.Debug.WriteLine("oldValue - " + oldValue);
            }
            catch { return; }
            Doc.Variables.Add("edocs_Page"+newPageText + "_" + VariablesText, oldValue);
            System.Diagnostics.Debug.WriteLine("newPageText + VariablesText" + newPageText + "_" + VariablesText);
        }
        public void setVarString(string varName, int page)
        {
            string endVar;
            if (varName == "page")
                try
                {
                    endVar = Doc.Variables["edocs_Page" + page + "_page"].Value;
                }
                catch
                {
                    Doc.Variables["edocs_Page" + page + "_page"].Value = "-";
                }
            try
            {
                endVar = Doc.Variables["edocs_Page" + Doc.Variables["edocs_Page" + page + "_page"].Value + "_" + varName].Value;
            }
            catch
            {
                Doc.Variables["edocs_Page" + Doc.Variables["edocs_Page" + page + "_page"].Value + "_" + varName].Value = "-";
            }
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
            saveLastRivision("last_page_rivision", MyRibbon.AutoRevString);
            saveLastRivision("last_date_rivision", MyRibbon.AutoDateString);
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
                        try
                        {
                            pageString = Doc.Variables["edocs_Page" + (pageNumberStrat + i) + "_page"].Value;
                            System.Diagnostics.Debug.WriteLine("pageString : " + pageString);
                            saveDocVariables("edocs_Page" + pageString, "rev", MyRibbon.AutoRevString);
                            saveDocVariables("edocs_Page" + pageString, "date", MyRibbon.AutoDateString);
                            setAllHeaderNumbers(pageNumberStrat + i);
                            System.Diagnostics.Debug.WriteLine("Changes Made in page : " + (pageNumberStrat + i));
                        }
                        catch
                        {
                            System.Diagnostics.Debug.WriteLine("Error in find pageString for page: " + pageNumberStrat + i);
                        }
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
        private void saveLastRivision(String var,String data)
        {
            try
            {
                Doc.Variables[var].Value = data;
            }
            catch
            {
                Doc.Variables.Add(var, data);
            }
    }
        private void setAllHeaderNumbers(int pageNumber)
        {
            try
            {
                //setFromFirstToIndex(pageString, pageNumber);
                if (toSaveUntilLastPage(Doc.Variables["edocs_Page" + (pageNumber) + "_page"].Value, pageNumber + 1))
                    setFromIndexToLast(Doc.Variables["edocs_Page" + (pageNumber) + "_page"].Value, pageNumber + 1, MyRibbon.AutoRevString, MyRibbon.AutoDateString, getFieldValue(pageNumber, "issue"), getFieldValue(pageNumber, "effective"));
            }
            catch
            {
                System.Diagnostics.Debug.WriteLine("Error in find pageString for page: " + pageNumber+1);
            }
        }
        private void setFromFirstToIndex(string pageString, int toPageNumber)
        {
            string pageTamplate = pageString.TrimEnd(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
            List<String> pagesArray = new List<String>();
            for (int i = 1; i < toPageNumber; i++)
            {
                string page = pageTamplate + i.ToString();
                pagesArray.Add(page);
                saveDocVariables("edocs_Page" + page, "rev", MyRibbon.AutoRevString);
                saveDocVariables("edocs_Page" + page, "date", MyRibbon.AutoDateString);
                System.Diagnostics.Debug.WriteLine("Changes Made in page : " + page);
            }
        }
        private void setFromIndexToLast(string pageString, int fromRealPageNumber,string rev,string date,string issue, string effective)
        {
            string pageTamplate = pageString.TrimEnd(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
            bool lastPage = false;
            int realPageIndex = fromRealPageNumber;
            while (!lastPage)
            { 
                try
                {
                    pageString = Doc.Variables["edocs_Page" + (realPageIndex) + "_page"].Value;
                    if (pageString.TrimEnd(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' }).Equals(pageTamplate))
                    {
                        string page = Doc.Variables["edocs_Page" + (realPageIndex) + "_page"].Value;
                        if (getFieldValue(realPageIndex, "rev") == null)
                        {
                            saveDocVariables("edocs_Page" + page, "rev", rev);
                            saveDocVariables("edocs_Page" + page, "date", date);
                            if (issue != null) saveDocVariables("edocs_Page" + page, "issue", issue);
                            if (effective != null) saveDocVariables("edocs_Page" + page, "effective", effective);
                            System.Diagnostics.Debug.WriteLine("Changes Made in page : " + page);
                        }
                        realPageIndex++;
                    }
                    else
                        lastPage = true;
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("Error in find pageString for page: " + realPageIndex);
                    lastPage = true;
                }
            }
        }
        private bool toSaveUntilLastPage(string pageString, int fromRealPageNumber)
        {
            string pageTamplate = pageString.TrimEnd(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
            for(int i=fromRealPageNumber; ;i++)
            {
                    pageString = Doc.Variables["edocs_Page" + (i) + "_page"].Value;
                    if (pageString.TrimEnd(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' }).Equals(pageTamplate))
                    {
                        if (getFieldValue(i, "rev")==null)
                            return true;
                    }
                    else
                        return false;
            }
        }
        private string getFieldValue(int pageNumber,string fieldName)
        {
            try
            {
                string endVar = Doc.Variables["edocs_Page" + Doc.Variables["edocs_Page" + (pageNumber) + "_page"].Value + "_" + fieldName].Value;
                return endVar;
            }
            catch
            {
                return null;
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
            for (int i = 1; i <= Doc.TablesOfContents.Count; i++)
                Doc.TablesOfContents[i].Range.Fields.Locked = -1;
            Globals.ThisAddIn.Application.ScreenRefresh();
            Globals.ThisAddIn.Application.Selection.Fields.Update();
            Doc.Fields.Update();
            for (int i = 1; i <= Doc.TablesOfContents.Count; i++)
                Doc.TablesOfContents[i].Range.Fields.Locked = 0;
        }
        private static int GetPageNumberOfRange(Word.Range range)
        {
            return (int)range.get_Information(Word.WdInformation.wdActiveEndAdjustedPageNumber);
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
            for (int t = 1; t <= Doc.TablesOfContents.Count; t++)
            {
                System.Diagnostics.Debug.WriteLine("TablesOfContents Count - " + Doc.TablesOfContents.Count);
                for (int i = 2; i <= Doc.TablesOfContents[t].Range.Paragraphs.Count; i++)
                {
                    if (IsAlert && settings.alert.worker.CancellationPending)
                        return;
                    Word.Paragraph pra = Doc.TablesOfContents[t].Range.Paragraphs[i];
                    System.Diagnostics.Debug.WriteLine("Paragraph Count - " + Doc.TablesOfContents[t].Range.Paragraphs.Count);
                    if (IsAlert && settings.alert.worker.CancellationPending)
                        return;
                    string rangeText = pra.Range.Text;
                    string[] numbers = Regex.Split(rangeText, @"\D+");
                    System.Diagnostics.Debug.WriteLine("Paragraph " + rangeText);

                    for (int z = numbers.Length - 1; z >= 0; z--)
                    {
                        if (IsAlert && settings.alert.worker.CancellationPending)
                            return;
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
                string pageTemplate = Doc.Variables["pageTemplate"].Value;
                switch (pageTemplate)
                {
                    case "X-P1":
                        pageValue = getFieldValue(pageNum, "X-P1");
                        break;
                    case "X-P-1":
                        pageValue = getFieldValue(pageNum, "X-P-1");
                        break;
                    case "X1":
                        pageValue = getFieldValue(pageNum, "X1");
                        break;
                    default:
                        pageValue = Doc.Variables["edocs_Page" + pageNum + "_page"].Value;
                        break;
                }
            }
            catch(Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
                return;
            }
            pra.Range.Find.ClearFormatting();
            object replaceText = pageValue;
                object replaceAll = Word.WdReplace.wdReplaceOne;
                object forward = false;
                object matchAllWord = false;
            if (pra.Range.Find.Execute(ref findText, ref missing, ref matchAllWord, ref missing, ref missing, ref missing, ref forward, Word.WdFindWrap.wdFindAsk, ref missing, ref replaceText, ref replaceAll, ref missing, true, ref missing, ref missing))
                    System.Diagnostics.Debug.WriteLine("FOUND - " + pageNum + " new page - " + pageValue);
                else
                    System.Diagnostics.Debug.WriteLine("Not FOUND");
            

        }
    }
}
