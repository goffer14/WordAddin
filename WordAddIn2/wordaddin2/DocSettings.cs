using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using WordAddIn2.Properties;
using System.Xml;
using System.Text;
using System.Security.Cryptography;

namespace WordAddIn2
{

    public class DocSettings
    {
        public Word.Document Doc;
        public object missing = System.Reflection.Missing.Value; 
        public string heading1_name;
        public string heading2_name;

        public string section1_Introduction‎1_name;
        public string section1_Introduction‎2_name;

        public HeadingPos Heading_pos1;
        public object Start;
        public string FooterText;
        public object End;
        public object CurrentPageNumber;
        public object NextPageNumber;
        public object What = Word.WdGoToItem.wdGoToPage;
        public object Which = Word.WdGoToDirection.wdGoToAbsolute;
        public object Miss = System.Reflection.Missing.Value;
        public bool IsAlert;

        public DocSettings(Word.Document ParentDoc)
        {
            Doc = ParentDoc;
            Heading_pos1 = new HeadingPos();
            setStylesToDoc();
            heading1_name = Doc.Styles[Doc.Variables["heading1_name"].Value].NameLocal;
            heading2_name = Doc.Styles[Doc.Variables["heading2_name"].Value].NameLocal;
            try
            {
                section1_Introduction‎1_name = ((Word.Style)Doc.Styles[Doc.Variables["introduction1"].Value]).NameLocal;
                section1_Introduction‎2_name = ((Word.Style)Doc.Styles[Doc.Variables["introduction2"].Value]).NameLocal;
            }
            catch { }

            get_datatable();

        }
        public bool dt_init()
        {
            // Unsecure eDoc for Edit.
            if (Doc.ProtectionType == Word.WdProtectionType.wdAllowOnlyReading)
            {
                unsecure_for_edit(Doc);
            }
            // Set a new DataBase if not stored one in the Document.
            foreach (Word.Variable varr in Doc.Variables)
            {
                if (varr.Name.Contains("edocs"))
                {
                    varr.Delete();
                }
            }
            Doc.Variables.Add("edocs_data_id", "is_edoc_good"); ;
            Heading_pos1 = new HeadingPos();
            int numofValInHeader = getTextPosInCell(1);
                    if (getTextPosInCell(2) + numofValInHeader != 5)
                    {
                        MessageBox.Show("Template problem - not eDoc template");
                        return false;
                    }                    
            return true;


        }
        public void setStylesToDoc()
        {
            try
            {
                string test1 = Doc.Variables["heading1_name"].Value;
            }
            catch
            {
                Doc.Variables.Add("heading1_name", "Heading 1");
            }
            try
            {
                string test1 = Doc.Variables["heading2_name"].Value;
            }
            catch
            {
                Doc.Variables.Add("heading2_name", "Heading 2");
            }
            try
            {
                string test1 = Doc.Variables["introduction1"].Value;
            }
            catch
            {
                Doc.Variables.Add("introduction1", "Introduction‎ 1");
            }
            try
            {
                string test1 = Doc.Variables["introduction2"].Value;
            }
            catch
            {
                Doc.Variables.Add("introduction2", "Introduction‎ 2");
            }
        }
        public bool insertSectionBreakBeforeHeading1()
        {
            Word.Range rng = Doc.Range();
            object heading1_name1 = heading1_name;
            rng.Find.set_Style(ref heading1_name1);
            rng.Find.Execute();
            int rangePageNumber = GetPageNumberOfRange(rng);
            if (rng.ListParagraphs[1].Range.Text.Trim() != "")
            {
                if (GetPageRange(rangePageNumber).Start != rng.Start && rangePageNumber!=1)
                {
                    Object beginPageNext = rng.Start;
                    Word.Range rangeForBreak = Doc.Range(ref beginPageNext, ref beginPageNext);
                    rangeForBreak.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                    rangePageNumber++;

                }
                if (rangePageNumber>1 && !(IsPageBreaks(GetPageRange(rangePageNumber- 1))))
                {
                    Object beginPageNext = rng.Start;
                    Word.Range rangeForBreak = Doc.Range(ref beginPageNext, ref beginPageNext);
                    rangeForBreak.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                }
                
            }
            int SectionCount = Doc.Sections.Count;
            if (SectionCount > 1)
            {
                
                    if (Doc.Sections[SectionCount].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious)
                    {
                    try
                    {
                        Doc.Sections[SectionCount].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
                        Doc.Sections[SectionCount].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                    }
                    catch (Exception ex)
                    {
                        Doc.Sections[SectionCount].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Select();
                        Word.Selection selection = Doc.ActiveWindow.Selection;
                        try
                        {
                            selection.HeaderFooter.LinkToPrevious = false;
                            Doc.Sections[SectionCount].Range.Select();
                            Doc.ActiveWindow.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                            Doc.ActiveWindow.Selection.Collapse();
                        }
                        catch (Exception ex1)
                        {
                        }
                    }
                        
                
                }
                try
                {
                    if (Doc.Sections[SectionCount].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Exists)
                        Doc.Sections[SectionCount].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Exists = false;
                }
                catch (Exception ex)
                { }
                try
                {
                    if (Doc.Sections[SectionCount].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Exists)
                        Doc.Sections[SectionCount].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                }
                catch (Exception ex)
                {}
                try
                {
                    if (Doc.Sections[SectionCount].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Exists)
                        Doc.Sections[SectionCount].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Exists = false;
                }
                catch (Exception ex)
                {  }

                return true;
            }
            return false;
        }
        public bool IsPageBreaks(Word.Range RangeToChack)
        {

            object findText = "^m"; // pageBreak
                if (RangeToChack.Find.Execute(findText))
                {
                    RangeToChack.Select();
                    RangeToChack.Delete();
                    RangeToChack.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                    return true;
                }
            
            findText = "^012"; // pageBreak
            if (RangeToChack.Find.Execute(findText))
            {
                RangeToChack.Select();
                RangeToChack.Delete();
                RangeToChack.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                return true;
            }
            findText = "^b";
            if (RangeToChack.Find.Execute(findText))
                return true;

            return false;
        }
        public int GetPageNumber(Word.Document doc)
        {
            return doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, System.Reflection.Missing.Value);
        }

        public void InesrtDatatoHeadingCells()
        {
            for (int i = 1; i <= Doc.Sections.Count; i++)
            {
                //build page cell
                object PageString = "\"edocs_PAGE_page\"";
                Word.Range range = getRangeFromCell(Doc.Sections[i], Heading_pos1.page_row, Heading_pos1.page_column, Heading_pos1.page_pos);
                range.Text = " ";
                range.End = range.Start;
                Doc.Application.Selection.Range.Fields.Add(range, Word.WdFieldType.wdFieldDocVariable, PageString, true);
                range.Start = range.Start + 25;
                Doc.Application.Selection.Range.Fields.Add(range, Word.WdFieldType.wdFieldEmpty, @"PAGE \* ARABIC", true);

                // build rev cell
                range = getRangeFromCell(Doc.Sections[i], Heading_pos1.rev_row, Heading_pos1.rev_column, Heading_pos1.rev_pos);
                BuildFields(range, "\"edocs_PAGE_rev\"", PageString);

                // build date cell
                range = getRangeFromCell(Doc.Sections[i], Heading_pos1.date_row, Heading_pos1.date_column, Heading_pos1.date_pos);
                BuildFields(range, "\"edocs_PAGE_date\"", PageString);


                if (Heading_pos1.text1_pos != -1)
                {
                    // build text1 cell
                    range = getRangeFromCell(Doc.Sections[i], Heading_pos1.text1_row, Heading_pos1.text1_column, Heading_pos1.text1_pos);
                    BuildFields(range, "\"edocs_PAGE_text1\"", PageString);
                }

                if (Heading_pos1.text2_pos != -1)
                {
                    // build text2 cell
                    range = getRangeFromCell(Doc.Sections[i], Heading_pos1.text2_row, Heading_pos1.text2_column, Heading_pos1.text2_pos);
                    BuildFields(range, "\"edocs_PAGE_text2\"", PageString);
                }
                if (Heading_pos1.text3_pos != -1)
                {
                    // build text3 cell
                    range = getRangeFromCell(Doc.Sections[i], Heading_pos1.text3_row, Heading_pos1.text3_column, Heading_pos1.text3_pos);
                    BuildFields(range, "\"edocs_PAGE_text3\"", PageString);
                }
                if (Heading_pos1.text4_pos != -1)
                {
                    // build text4 cell
                    range = getRangeFromCell(Doc.Sections[i], Heading_pos1.text4_row, Heading_pos1.text4_column, Heading_pos1.text4_pos);
                    BuildFields(range, "\"edocs_PAGE_text4\"", PageString);
                }
            }
        }
        public void BuildFields(Word.Range range, object textString, object PageString)
        {
            range.Text = " ";
            range.End = range.Start;
            Doc.Application.Selection.Range.Fields.Add(range, Word.WdFieldType.wdFieldDocVariable, textString, true);
            range.Start = range.Start + 25;
            Doc.Application.Selection.Range.Fields.Add(range, Word.WdFieldType.wdFieldDocVariable, PageString, true);
            range.Start = range.Start + 25;
            Doc.Application.Selection.Range.Fields.Add(range, Word.WdFieldType.wdFieldEmpty, @"PAGE \* ARABIC", true);
        }
        public Word.Range GetPageRange(int PageNubmer)
        {
            CurrentPageNumber = (Convert.ToInt32(PageNubmer.ToString()));
            NextPageNumber = (Convert.ToInt32((PageNubmer + 1).ToString()));
            // Get start position of current page
            Start = Doc.GoTo(ref What, ref Which, ref CurrentPageNumber, ref Miss).Start;
       //     if (Doc.TablesOfContents.Count > 0 && PageNubmer == Doc.TablesOfContents[1].Range.Paragraphs[Doc.TablesOfContents[1].Range.Paragraphs.Count].Range.Information[Word.WdInformation.wdActiveEndPageNumber] && PageNubmer != Doc.TablesOfContents[1].Range.Paragraphs[2].Range.Information[Word.WdInformation.wdActiveEndPageNumber])
         //       Start = Doc.TablesOfContents[1].Range.End;
               // Get end position of current page                                
               End = Doc.GoTo(ref What, ref Which, ref NextPageNumber, ref Miss).End;
            // Get text
            if (Convert.ToInt32(Start.ToString()) != Convert.ToInt32(End.ToString()))
                return Doc.Range(ref Start, ref End);
            else
                return Doc.Range(ref Start);
        }
        public Dictionary<string, string> saveAllRangsText()
        {

            Dictionary<string, string>  original_sections = new Dictionary<string, string> { }; ;
            string edoc_page_text;
            string hash_string;
            // Populate "original_sections" Dictionary with hashes from the Document
            // Keys go in format ["chapter-section"]
            for (int i = GetFirstHeader()[1]; i <= GetPageNumber(Doc); i++)
            {
                if (settings.alert.worker.CancellationPending)
                    return null;
                Word.Range rng = GetPageRange(i);
                edoc_page_text = Doc.Variables["edocs_PAGE" + i + "_page"].Value;
                hash_string = GetHashString(rng.Text.Trim());
                original_sections.Add(edoc_page_text, hash_string);
            }
            return original_sections;

        }
        public int[] GetFirstHeader()
        {
            object styleTo;
            try
            {
                styleTo = Doc.Styles[section1_Introduction‎1_name].NameLocal;
            }
            catch
            {
                styleTo = Doc.Styles[heading1_name].NameLocal;
                return findByStyle(Doc.Range(), styleTo);
            }
            int[] pages = findByStyle(Doc.Range(), styleTo);
            if (pages != null)
                return pages;
            styleTo = Doc.Styles[heading1_name].NameLocal;
            return findByStyle(Doc.Range(), styleTo);
        }
        public static int[] findByStyle(Word.Range rng, object styleTo)
        {
            rng.Find.set_Style(ref styleTo);
            if (rng.Find.Execute())
            {
                try
                {
                    return new int[] { 0, GetPageNumberOfRange(rng) };
                }
                catch { }
            }
            else
                return null;
            return new int[] { 1, 1 };
        }
        public static byte[] GetHash(string inputString)
        {
            HashAlgorithm algorithm = MD5.Create();  //or use SHA1.Create();
            return algorithm.ComputeHash(Encoding.UTF8.GetBytes(inputString));
        }
        public string GetHashString(string inputString)
        {
            StringBuilder sb = new StringBuilder();
            foreach (byte b in GetHash(inputString))
                sb.Append(b.ToString("X2"));

            return sb.ToString();
        }
        private void get_datatable() {
            try
            {
                //h1
                Heading_pos1.H1_column = Int32.Parse(Doc.Variables["edocs_H1_column"].Value);
                Heading_pos1.H1_row = Int32.Parse(Doc.Variables["edocs_H1_row"].Value);
                Heading_pos1.H1_pos = Int32.Parse(Doc.Variables["edocs_H1_pos"].Value);

                //h2
                Heading_pos1.H2_column = Int32.Parse(Doc.Variables["edocs_H2_column"].Value);
                Heading_pos1.H2_row = Int32.Parse(Doc.Variables["edocs_H2_row"].Value);
                Heading_pos1.H2_pos = Int32.Parse(Doc.Variables["edocs_H2_pos"].Value);

                //rev
                Heading_pos1.rev_column = Int32.Parse(Doc.Variables["edocs_rev_column"].Value);
                Heading_pos1.rev_row = Int32.Parse(Doc.Variables["edocs_rev_row"].Value);
                Heading_pos1.rev_pos = Int32.Parse(Doc.Variables["edocs_rev_pos"].Value);
                //date
                Heading_pos1.date_column = Int32.Parse(Doc.Variables["edocs_date_column"].Value);
                Heading_pos1.date_row = Int32.Parse(Doc.Variables["edocs_date_row"].Value);
                Heading_pos1.date_pos = Int32.Parse(Doc.Variables["edocs_date_pos"].Value);

                //page
                Heading_pos1.page_column = Int32.Parse(Doc.Variables["edocs_page_column"].Value);
                Heading_pos1.page_row = Int32.Parse(Doc.Variables["edocs_page_row"].Value);
                Heading_pos1.page_pos = Int32.Parse(Doc.Variables["edocs_page_pos"].Value);
            }
            catch { }
            try
            {
                //text1
                Heading_pos1.text1_column = Int32.Parse(Doc.Variables["edocs_text1_column"].Value);
                Heading_pos1.text1_row = Int32.Parse(Doc.Variables["edocs_text1_row"].Value);
                Heading_pos1.text1_pos = Int32.Parse(Doc.Variables["edocs_text1_pos"].Value);
            }
            catch { }
            try
            {
                //text2
                Heading_pos1.text2_column = Int32.Parse(Doc.Variables["edocs_text2_column"].Value);
                Heading_pos1.text2_row = Int32.Parse(Doc.Variables["edocs_text2_row"].Value);
                Heading_pos1.text2_pos = Int32.Parse(Doc.Variables["edocs_text2_pos"].Value);
            }
            catch { }
            try
            {
                //text3
                Heading_pos1.text3_column = Int32.Parse(Doc.Variables["edocs_text3_column"].Value);
                Heading_pos1.text3_row = Int32.Parse(Doc.Variables["edocs_text3_row"].Value);
                Heading_pos1.text3_pos = Int32.Parse(Doc.Variables["edocs_text3_pos"].Value);
            }
            catch { }
            try
            {
                //text4
                Heading_pos1.text4_column = Int32.Parse(Doc.Variables["edocs_text4_column"].Value);
                Heading_pos1.text4_row = Int32.Parse(Doc.Variables["edocs_text4_row"].Value);
                Heading_pos1.text4_pos = Int32.Parse(Doc.Variables["edocs_text4_pos"].Value);
            }
            catch { }
        }
        public bool unsecure_for_edit(Word.Document DocToUnSecure)
        {
            object password = settings.doc_password;
                try {
                DocToUnSecure.Unprotect(ref password);
                }
                catch {
                    return false; 
                }
            return true;
        }
        public void insertRev_Rdate(string PageString, string rev, string r_date)
        {
            string page_text = "edocs_PAGE" + PageString;
            try
            {
                Doc.Variables[page_text + "_rev"].Delete();
            }
            catch { }
            try
            {
                Doc.Variables[page_text + "_date"].Delete();
            }
            catch { }
            Doc.Variables.Add(page_text + "_rev", rev);
            Doc.Variables.Add(page_text + "_date", r_date);

        }
        public void insertTextToData(string PageString, string text1, string text2, string text3, string text4)
        {
            if (text1 == "")
                text1 = " ";
            if (text2 == "")
                text2 = " ";
            if (text3 == "")
                text3 = " ";
            if (text4 == "")
                text4 = " ";
            string page_text = "edocs_PAGE" + PageString;
            if (Heading_pos1.text1_pos != -1)
            {
                try
                {
                    Doc.Variables[page_text + "_text1"].Delete();
                }
                catch { }
                Doc.Variables.Add(page_text + "_text1", text1);
            }
            if (Heading_pos1.text2_pos != -1)
            {
                try
                {
                    Doc.Variables[page_text + "_text2"].Delete();
                }
                catch { }
                Doc.Variables.Add(page_text + "_text2", text2);
            }
            if (Heading_pos1.text3_pos != -1)
            {
                try
                {
                    Doc.Variables[page_text + "_text3"].Delete();
                }
                catch { }
                Doc.Variables.Add(page_text + "_text3", text3);
            }
            if (Heading_pos1.text4_pos != -1)
            {
                try
                {
                    Doc.Variables[page_text + "_text4"].Delete();
                }
                catch { }
                Doc.Variables.Add(page_text + "_text4", text4);
            }


        }

        public void ChangeRev (string[] formPage ,int fromPage,int ToPage, string rev , string r_date)
       {
           for (int i = fromPage; i <= ToPage; i++)
               insertRev_Rdate(formPage[i], rev, r_date);
           
        }
        public void ChangeText(string[] formPage, int fromPage, int ToPage, string text1, string text2,string text3,string text4)
        {
            for (int i = fromPage; i <= ToPage; i++)
                insertTextToData(formPage[i], text1, text2, text3, text4);

        }
        public bool PageNumberFromHeaders(int DocPageNumber)
        {

            Word.Range rng = Doc.Range();
            List<header> PageHeader = new List<header>();
            int firstHeaderPage = 0;
            string H1_string = "null";
            string H2_string = "null";
            object heading2_name1 = heading2_name;
            rng.Find.set_Style(ref heading2_name1);
            int middleHeader = 0;
            bool onceNotTwice=true;
            int rngPageNumber = 0;
            int StartOfH1 = 0;
            
                for (int i = 1; i <= Doc.Sections.Count; i++)
                {
                    try
                    {
                        Doc.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = false;
                    }
                    catch { }
                    try
                    {
                        Doc.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = false;
                    }
                    catch { }
                }

            
            try
            {
                if (Doc.Sections.Count > 1)
                {
                    int[] arr1 = GetFirstHeader();
                    if (arr1[0] == 0)
                        PageHeader.Add(new header(arr1[1], "", ""));
                }
                while (rng.Find.Execute())
                {

                    if (IsAlert && settings.alert.worker.CancellationPending)
                        return false;
                    rngPageNumber = GetPageNumberOfRange(rng);
                    if (onceNotTwice)
                    {
                        StartOfH1 = rngPageNumber-1;
                        onceNotTwice = false;
                    }

                        if (rng.ListParagraphs[1].Range.Text.Trim() != "" && rng.Text.Length > 3 && firstHeaderPage < rngPageNumber)
                    {
                        if(middleHeader!=0&& middleHeader< rngPageNumber)
                        {
                            PageHeader.Add(new header(middleHeader, H1_string, H2_string));
                            middleHeader = 0;
                        }
                        H2_string = rng.ListParagraphs[1].Range.ListFormat.ListValue.ToString();
                        H1_string = rng.ListParagraphs[1].Range.ListFormat.ListString.ToString();
                        if (H1_string != "" && H1_string.IndexOf(".", (H1_string.Length - 1)) > 0)
                            H1_string = H1_string.Remove(H1_string.Length - 1, 1);
                        H1_string = H1_string.Substring(0, H1_string.LastIndexOf("."));
                        firstHeaderPage = rngPageNumber;
                        PageHeader.Add(new header(firstHeaderPage, H1_string, H2_string));
                        Doc.Application.Selection.GoTo(ref What, ref Which, firstHeaderPage, ref missing);
                    }
                    else if(firstHeaderPage == rngPageNumber && rng.ListParagraphs[1].Range.Text.Trim() != "" && rng.Text.Length > 3)
                    {
                        H2_string = rng.ListParagraphs[1].Range.ListFormat.ListValue.ToString();
                        H1_string = rng.ListParagraphs[1].Range.ListFormat.ListString.ToString();
                        if (H1_string != "" && H1_string.IndexOf(".", (H1_string.Length - 1)) > 0)
                            H1_string = H1_string.Remove(H1_string.Length - 1, 1);
                        H1_string = H1_string.Substring(0, H1_string.LastIndexOf("."));
                        middleHeader = rngPageNumber+1;
                    }
                    rng.Start = rng.End;
                }
                if(middleHeader== DocPageNumber)
                {
                    PageHeader.Add(new header(middleHeader, H1_string, H2_string));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something went wrong: " + firstHeaderPage + Environment.NewLine + " Error Msg: " + ex.Message.ToString());
                if (IsAlert)
                    settings.alert.Close();
                return false;
            }
            return InsertPageNumberToDataBase(PageHeader,DocPageNumber, StartOfH1);
        }
        public bool InsertPageNumberToDataBase(List<header> PageHeader,int DocPageNumber,int StartOfH1)
        {
            string valtext;
            try
            {
                for (int i = 0; i < PageHeader.Count-1; i++)
                {
                    for (int j = 1; j <= PageHeader[i + 1].pageNum - PageHeader[i].pageNum; j++)
                    {
                        if (IsAlert && settings.alert.worker.CancellationPending)
                            return false;
                        string edoc_page_text = "edocs_PAGE" + ((PageHeader[i].pageNum) + (j - 1)) + "_page";
                        if ((PageHeader[i].pageNum) + (j - 1) > StartOfH1)
                            valtext = PageHeader[i].H1_num + " - " + PageHeader[i].H2_num + " - P-" + j.ToString();
                        else
                            valtext = "INTR-" +j.ToString();
                        try
                        {
                            Doc.Variables[edoc_page_text].Delete();
                        }
                        catch
                        { }
                        if (settings.save_text)
                            changePageData(((PageHeader[i].pageNum) + (j - 1)), valtext);
                        Doc.Variables.Add(edoc_page_text, valtext);
                    }
                }
                int lastPage = PageHeader.Count - 1;
                for (int j = 0; j <= DocPageNumber - PageHeader[lastPage].pageNum; j++)
                {
                    if (IsAlert && settings.alert.worker.CancellationPending)
                        return false;
                    string page_text = "edocs_PAGE" + ((PageHeader[lastPage].pageNum) + j) + "_page";
                    valtext = PageHeader[lastPage].H1_num + " - " + PageHeader[lastPage].H2_num + " - P-" + (j+1).ToString();
                    try
                    {
                        Doc.Variables[page_text].Delete();
                    }
                    catch
                    { }
                    if (settings.save_text)
                        changePageData(((PageHeader[lastPage].pageNum) + j), valtext);
                    Doc.Variables.Add(page_text, valtext);
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
        public void changePageData(int RangePage,string valtext)
        {
            string original_section_text = "";
            Word.Range rng = GetPageRange(RangePage);

            string section_text = GetHashString(rng.Text.Trim());
            settings.original_sections.TryGetValue(valtext, out original_section_text);
            if (section_text != original_section_text)
            {
                try
                {
                    Doc.Variables["edocs_PAGE" + valtext + "_rev"].Delete();
                }
                catch
                { }
                try
                {
                    Doc.Variables["edocs_PAGE" + valtext + "_date"].Delete();
                }
                catch
                { }
                Doc.Variables.Add("edocs_PAGE" + valtext + "_rev", MyRibbon.AutoRevString);
                Doc.Variables.Add("edocs_PAGE" + valtext + "_date", MyRibbon.AutoDateString);
                settings.original_sections.Remove(valtext);
                settings.original_sections.Add(valtext, section_text);
            }
        }
        public void InsertCodeToHeader()
        {
            for (int i = 1; i <= Doc.Sections.Count; i++)
            {
                Word.Range range = getRangeFromCell(Doc.Sections[i], Heading_pos1.page_row, Heading_pos1.page_column, Heading_pos1.page_pos);
                range.Text = " ";
                range = getRangeFromCell(Doc.Sections[i], Heading_pos1.date_row, Heading_pos1.date_column, Heading_pos1.date_pos);
                range.Text = " ";

                Word.Range tableRange = getRangeFromCell(Doc.Sections[i], Heading_pos1.H1_row, Heading_pos1.H1_column, Heading_pos1.H1_pos);
                tableRange.Text = " ";
                object txtstyleRef = "\"" + heading1_name + "\" \\n";
                tableRange.End = tableRange.Start;
                Doc.Application.Selection.Range.Fields.Add(tableRange, Word.WdFieldType.wdFieldStyleRef, txtstyleRef, true);


                tableRange = getRangeFromCell(Doc.Sections[i], Heading_pos1.H1_row, Heading_pos1.H1_column, Heading_pos1.H1_pos);
                tableRange.End = tableRange.End - 1;
                tableRange.Start = tableRange.End;
                txtstyleRef = "\"" + heading1_name + "\"";
                Doc.Application.Selection.Range.Fields.Add(tableRange, Word.WdFieldType.wdFieldStyleRef, txtstyleRef, true);

                tableRange = getRangeFromCell(Doc.Sections[i], Heading_pos1.H2_row, Heading_pos1.H2_column, Heading_pos1.H2_pos);
                tableRange.Text = " ";
                txtstyleRef = "\"" + heading2_name + "\" \\n";
                tableRange.End = tableRange.Start;
                Doc.Application.Selection.Range.Fields.Add(tableRange, Word.WdFieldType.wdFieldStyleRef, txtstyleRef, true);

                tableRange = getRangeFromCell(Doc.Sections[i], Heading_pos1.H2_row, Heading_pos1.H2_column, Heading_pos1.H2_pos);
                tableRange.End = tableRange.End - 1;
                tableRange.Start = tableRange.End;
                txtstyleRef = "\"" + heading2_name + "\"";
                Doc.Application.Selection.Range.Fields.Add(tableRange, Word.WdFieldType.wdFieldStyleRef, txtstyleRef, true);
            }
        }
            


        public bool isIntr()
        {
            Word.Range rng = Doc.Range();
            object section1_Introduction‎1_name1 = section1_Introduction‎1_name;
            rng.Find.set_Style(ref section1_Introduction‎1_name1);
            if (rng.Find.Execute())
                return true;
                return false;

        }
        public void resetHeadersInFirstSection()
        {

            Word.Range tableRange = getRangeFromCell(Doc.Sections[1], Heading_pos1.H1_row, Heading_pos1.H1_column, Heading_pos1.H1_pos);
            tableRange.Text = " ";
            object txtstyleRef = "\"" + section1_Introduction‎1_name + "\"";
            tableRange.End = tableRange.Start;
            Doc.Application.Selection.Range.Fields.Add(tableRange, Word.WdFieldType.wdFieldStyleRef, txtstyleRef, true);


            tableRange = getRangeFromCell(Doc.Sections[1], Heading_pos1.H2_row, Heading_pos1.H2_column, Heading_pos1.H2_pos);
            tableRange.Text = " ";
            txtstyleRef = "\"" + section1_Introduction‎2_name + "\"";
            tableRange.End = tableRange.Start;
            Doc.Application.Selection.Range.Fields.Add(tableRange, Word.WdFieldType.wdFieldStyleRef, txtstyleRef, true);
        }
        public void UpDateFields()
        {
          //  Globals.ThisAddIn.Application.Selection.Fields.Update();
            Globals.ThisAddIn.Application.ActiveDocument.Fields.Update();
            Globals.ThisAddIn.Application.ScreenRefresh();
        }
        public void arrangeHeader1()
        {
            Word.Range rng = Doc.Range();
            object heading1_name1 = heading1_name;
            rng.Find.set_Style(ref heading1_name1);
            while (rng.Find.Execute())
            {
                if (IsAlert && settings.alert.worker.CancellationPending)
                    return;
                if (rng.ListParagraphs[1].Range.Text.Trim() != "" && rng.Start != GetPageRange(GetPageNumberOfRange(rng)).Start)
                {
                    Object beginPageNext = rng.Start;
                    Word.Range rangeForBreak = Doc.Range(ref beginPageNext, ref beginPageNext);
                    rangeForBreak.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
                rng.SetRange(rng.End, Doc.Range().End);
            }
        }

        public class CustomCell
        {
            public int cellRow;
            public int cellColumn;
            public int cellRowSpan;
            public int cellColumnSpan;
            public String cellText;
        }
        public class textPos
        {
            public int cellRow;
            public int cellColumn;
            public int cellPos;
        }
        public int getTextPosInCell(int searchInPos)
        {
            Word.Table wordTable = Doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Tables[1];
            String s = Doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Tables[1].Range.XML;

            if (searchInPos == 2)
            {
                try {
                    wordTable = Doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Tables[1];
                    s = Doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Tables[1].Range.XML;
                }
                catch
                {
                    return 0;
                }
            }
            Word.Cell wordCell = wordTable.Cell(1, 1);
            
            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
            xmlDoc.LoadXml(s);
            int numOfFind = 0;
            System.Xml.XmlNamespaceManager nsmgr = new System.Xml.XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", "http://schemas.microsoft.com/office/word/2003/wordml");
            while (wordCell != null)
            {
                CustomCell cell = new CustomCell();
                cell.cellRow = wordCell.RowIndex;
                cell.cellColumn = wordCell.ColumnIndex;
                int colspan;
                System.Xml.XmlNode exactCell = xmlDoc.SelectNodes("//w:tr[" + wordCell.RowIndex.ToString() + "]/w:tc[" + wordCell.ColumnIndex.ToString() + "]/w:tcPr/w:gridSpan", nsmgr)[0];
                if (exactCell != null)
                {
                    colspan = Convert.ToInt16(exactCell.Attributes["w:val"].Value);
                }
                else
                {
                    colspan = 1;
                }

                int rowspan = 1;
                Boolean endRows = false;
                int nextRows = wordCell.RowIndex + 1;
                System.Xml.XmlNode exactCellVMerge = xmlDoc.SelectNodes("//w:tr[" + wordCell.RowIndex.ToString() + "]/w:tc[" + wordCell.ColumnIndex.ToString() + "]/w:tcPr/w:vmerge", nsmgr)[0];

                if ((exactCellVMerge == null) || (exactCellVMerge != null && exactCellVMerge.Attributes["w:val"] == null))
                {
                    rowspan = 1;
                }
                else
                {
                    while (nextRows <= wordTable.Rows.Count && !endRows)
                    {
                        System.Xml.XmlNode nextCellMerge = xmlDoc.SelectNodes("//w:tr[" + nextRows.ToString() + "]/w:tc[" + wordCell.ColumnIndex.ToString() + "]/w:tcPr/w:vmerge", nsmgr)[0];
                        if (nextCellMerge != null && (nextCellMerge.Attributes["w:val"] == null))
                        {
                            nextRows++;
                            rowspan++;
                            continue;
                        }
                        else
                        {
                            endRows = true;
                        }
                    }
                }

                if (wordCell.Range.Text.Contains("-Heading 1-"))
                {
                    Doc.Variables.Add("edocs_H1_column", cell.cellColumn);
                    Doc.Variables.Add("edocs_H1_row", cell.cellRow);
                    Doc.Variables.Add("edocs_H1_pos", searchInPos);

                    Heading_pos1.H1_column = cell.cellColumn;
                    Heading_pos1.H1_row = cell.cellRow;
                    Heading_pos1.H1_pos = searchInPos;
                    numOfFind++;
                  //  wordCell.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                }
                else if (wordCell.Range.Text.Contains("-Heading 2-"))
                {
                    Doc.Variables.Add("edocs_H2_column", cell.cellColumn);
                    Doc.Variables.Add("edocs_H2_row", cell.cellRow);
                    Doc.Variables.Add("edocs_H2_pos", searchInPos);

                    Heading_pos1.H2_column = cell.cellColumn;
                    Heading_pos1.H2_row = cell.cellRow;
                    Heading_pos1.H2_pos = searchInPos;
                    numOfFind++;
                 //   wordCell.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                }
                else if(wordCell.Range.Text.Contains("-Revision-"))
                {
                    Doc.Variables.Add("edocs_rev_column", cell.cellColumn);
                    Doc.Variables.Add("edocs_rev_row", cell.cellRow);
                    Doc.Variables.Add("edocs_rev_pos", searchInPos);
                    Heading_pos1.rev_column = cell.cellColumn;
                    Heading_pos1.rev_row = cell.cellRow;
                    Heading_pos1.rev_pos = searchInPos;
                    numOfFind++;
                //    wordCell.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                }
                else if(wordCell.Range.Text.Contains("-Date-"))
                {
                    Doc.Variables.Add("edocs_date_column", cell.cellColumn);
                    Doc.Variables.Add("edocs_date_row", cell.cellRow);
                    Doc.Variables.Add("edocs_date_pos", searchInPos);
                    Heading_pos1.date_column = cell.cellColumn;
                    Heading_pos1.date_row = cell.cellRow;
                    Heading_pos1.date_pos = searchInPos;
                    numOfFind++;
                //    wordCell.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                }
                else if(wordCell.Range.Text.Contains("-Page-"))
                {
                    Doc.Variables.Add("edocs_page_column", cell.cellColumn);
                    Doc.Variables.Add("edocs_page_row", cell.cellRow);
                    Doc.Variables.Add("edocs_page_pos", searchInPos);
                    Heading_pos1.page_column = cell.cellColumn;
                    Heading_pos1.page_row = cell.cellRow;
                    Heading_pos1.page_pos = searchInPos;
                    numOfFind++;
                 //   wordCell.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                }
                else if (wordCell.Range.Text.Contains("-Text1-"))
                {
                    Doc.Variables.Add("edocs_text1_column", cell.cellColumn);
                    Doc.Variables.Add("edocs_text1_row", cell.cellRow);
                    Doc.Variables.Add("edocs_text1_pos", searchInPos);
                    Heading_pos1.text1_column = cell.cellColumn;
                    Heading_pos1.text1_row = cell.cellRow;
                    Heading_pos1.text1_pos = searchInPos;
                    //   wordCell.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                }
                else if (wordCell.Range.Text.Contains("-Text2-"))
                {
                    Doc.Variables.Add("edocs_text2_column", cell.cellColumn);
                    Doc.Variables.Add("edocs_text2_row", cell.cellRow);
                    Doc.Variables.Add("edocs_text2_pos", searchInPos);
                    Heading_pos1.text2_column = cell.cellColumn;
                    Heading_pos1.text2_row = cell.cellRow;
                    Heading_pos1.text2_pos = searchInPos;
                    //   wordCell.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                }
                else if (wordCell.Range.Text.Contains("-Text3-"))
                {
                    Doc.Variables.Add("edocs_text3_column", cell.cellColumn);
                    Doc.Variables.Add("edocs_text3_row", cell.cellRow);
                    Doc.Variables.Add("edocs_text3_pos", searchInPos);
                    Heading_pos1.text3_column = cell.cellColumn;
                    Heading_pos1.text3_row = cell.cellRow;
                    Heading_pos1.text3_pos = searchInPos;
                    //   wordCell.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                }
                else if (wordCell.Range.Text.Contains("-Text4-"))
                {
                    Doc.Variables.Add("edocs_text4_column", cell.cellColumn);
                    Doc.Variables.Add("edocs_text4_row", cell.cellRow);
                    Doc.Variables.Add("edocs_text4_pos", searchInPos);
                    Heading_pos1.text4_column = cell.cellColumn;
                    Heading_pos1.text4_row = cell.cellRow;
                    Heading_pos1.text4_pos = searchInPos;
                    //   wordCell.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                }

                wordCell = wordCell.Next;
            }
                return numOfFind;
        }
        public static Word.Range getRangeFromCell(Word.Section section, int header_row , int header_column , int header_pos)
        {
            Word.HeaderFooter header_footer = null;
            if (header_pos == 1)
                header_footer = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            else
                header_footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            Word.Range headerRange = header_footer.Range;
            Word.Table tbl = headerRange.Tables[1];
           // tbl.Cell(Heading.Header_row, Heading.Header_column).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly;
            return tbl.Cell(header_row, header_column).Range;
        }
        private static int GetPageNumberOfRange(Word.Range range)
        {
            return (int)range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
        }
    }
}