using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Text.RegularExpressions;
//using Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;
using HtmlAgilityPack;
using System.Data.OleDb;
using System.Windows.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
//using Document = Microsoft.Office.Interop.Word.Document;
using System.Configuration;
using Microsoft.Office.Interop.Word;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace TemplateConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();

            Globals.selectedDocument = "";

            treeView1.SetValue(VirtualizingStackPanel.IsVirtualizingProperty, true);
            treeView1.SetValue(VirtualizingStackPanel.VirtualizationModeProperty, VirtualizationMode.Recycling);

        }

        private TreeNode LoadDirectory(DirectoryInfo di)
        {
            if (!di.Exists)
                return null;

            TreeNode output = new TreeNode(di.Name, 0, 0);

            foreach (var subDir in di.GetDirectories())
            {
                try
                {
                    output.Nodes.Add(LoadDirectory(subDir));
                }
                catch (IOException ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
                catch { }
            }

            foreach (var file in di.GetFiles())
            {
                if (file.Exists)
                {
                    output.Nodes.Add(file.Name);
                }
            }

            return output;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var browser = new System.Windows.Forms.FolderBrowserDialog();
            //browser.RootFolder = Environment.SpecialFolder.MyDocuments;
            //browser.SelectedPath = 
            System.Windows.Forms.DialogResult result = browser.ShowDialog();

            string tempPath = "";

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                tempPath = browser.SelectedPath; // prints path

                Globals.directoryInfo = new DirectoryInfo(tempPath);

                if (Globals.directoryInfo.Exists)
                {
                    try
                    {
                        treeView1.Items.Clear();
                        treeView1.Items.Add(CreateDirectoryNode(Globals.directoryInfo));
                    }
                    catch (Exception ex)
                    {
                        MessageBoxResult exception = System.Windows.MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private static TreeViewItem CreateDirectoryNode(DirectoryInfo directoryInfo)
        {
            var directoryNode = new TreeViewItem { Header = directoryInfo.Name };
            foreach (var directory in directoryInfo.GetDirectories())
                directoryNode.Items.Add(CreateDirectoryNode(directory));

            foreach (var file in directoryInfo.GetFiles())
                directoryNode.Items.Add(new TreeViewItem { Header = file.Name });

            return directoryNode;

        }

        private void treeView1_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            label1.Content = "Initiating MS Word.";
            label1.Refresh();

            if (treeView1.Items.Count >= 0)
            {
                var tree = sender as System.Windows.Controls.TreeView;

                if (tree.SelectedItem is TreeViewItem)
                {
                    // ... Handle a TreeViewItem.
                    var item = tree.SelectedItem as TreeViewItem;
                    Globals.selectedDocument = item.Header.ToString();
                }
                else if (tree.SelectedItem is string)
                {
                    // ... Handle a string.
                    Globals.selectedDocument = tree.SelectedItem.ToString();
                }
            }

            string document = System.IO.Path.Combine(Globals.directoryInfo.FullName, Globals.selectedDocument);
            //MessageBoxResult result = System.Windows.MessageBox.Show(myMessage);

            label1.Content = "Reading: " + Globals.selectedDocument + " ... please wait.";
            label1.Refresh();

            string htmlText = "";

            try
            {
                //htmlText = await renderHTML.SearchAndHighlight(document);
                //htmlText = await renderHTML.SearchAndHighlightXMLDoc(document);
                htmlText = SearchAndHighlightXMLDoc(document);
                btnConvert.IsEnabled = true;
            } 
            catch (Exception ex)
            {
                htmlText = "<HTML><BODY><H1> Unable to display document! </H1> <br />" + ex.Message + "</BODY><?HTML>";
            }

            // load lookup
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(ConfigurationManager.AppSettings["xmlLookup"]);
            XmlDocument xmlDoc2 = new XmlDocument();
            xmlDoc2.Load(ConfigurationManager.AppSettings["xmlLookupLinkTable"]);

            htmlOutput.NavigateToString(htmlText);
            label1.Content = Globals.selectedDocument;

            // find all variables identified in our displayed html (where div class is 'field')
            var divFieldsXPath = "//div[contains(@class,'field')]";

            // reset list of mapped fields
            Globals.mergeFieldMapping.Clear();

            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(htmlText);

            HtmlNodeCollection listOfFields = htmlDoc.DocumentNode.SelectNodes(divFieldsXPath);
            // check there are actually some fields to iterate round

            if (listOfFields != null)
            {
                foreach (var node in listOfFields)
                {
                    // find the Merge Field Mapping
                    Regex fieldName = new Regex("[A-Z][A-Z]?[A-Z][0-9]."); // will find TF09 and ABC09

                    Match m = fieldName.Match(node.InnerText, 0);

                    string mappedMergeField = "unmapped";
                    // search xml merged field table with this xpath to get match "//Auto-Match-Screen-Fields/Data-Collection-Field-Name[../Field-Code-Lookup='ESD03']/text()"

                    XmlNode mappedNode = xmlDoc.SelectSingleNode("//Auto-Match-Screen-Fields/Merge-Field-Name[../Field-Code-Lookup='" + m.Value + "']/text()");
                    if (mappedNode != null)
                    {
                        mappedMergeField = mappedNode.InnerText;
                    }

                    // try the linked table
                    XmlNode mappedNodeLink = xmlDoc2.SelectSingleNode("//LINK_TABLE/MERGED-FIELD-NAME[../FIELD-CODE='" + m.Value + "']/text()");
                    if (mappedNodeLink != null)
                    {
                        mappedMergeField = mappedNodeLink.InnerText;
                    }

                    Globals.mergeFieldMapping.Add(new mapMergeField() { solcaseField = m.Value, mergeField = mappedMergeField });
                }
            }

            // Bind list to data grid
            dgMap.ItemsSource = Globals.mergeFieldMapping;
            CollectionViewSource.GetDefaultView(dgMap.ItemsSource).Refresh();
            dgMap.UpdateLayout();
            dgMap.Refresh();

        }

        public string SearchAndHighlightXMLDoc(string document)
        {
            // this method uses Open XML to read the docx version of the template

            //Microsoft.Office.Interop.Word.Application Word97 = new Microsoft.Office.Interop.Word.Application();
            //Word97.WordBasic.DisableAutoMacros();

            using (WordprocessingDocument wordDocx =
                WordprocessingDocument.Open(document, true))
            {

                // Assign a reference to the existing document body.
                //Body body = wordDoc.MainDocumentPart.Document.Body;
                // create merge field like «Estate_Details_Salutation_ESD01»
                //string wordMergeFieldTxt = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", item.mergeField);
                //SimpleField simpleField1 = new SimpleField() { Instruction = wordMergeFieldTxt };

                string docText = null;

                using (StreamReader sr = new StreamReader(wordDocx.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexIfCondition = new Regex(@"\[&amp;If.[A-Z]*[0-9]*[=,<>].*?\]", RegexOptions.IgnoreCase);
                Regex regexElseCondition = new Regex(@"\[&amp;Else.?\]", RegexOptions.IgnoreCase);
                Regex regexEndIfCondition = new Regex(@"\[&amp;EndIf.*?\]", RegexOptions.IgnoreCase);
                Regex regexForEach = new Regex(@"\[&amp;FOREACH.*?\]", RegexOptions.IgnoreCase);
                Regex regexEndFor = new Regex(@"\[&amp;ENDFOR.*?\]", RegexOptions.IgnoreCase);
                Regex regexIncludes = new Regex(@"\[&amp;Include.*?\]", RegexOptions.IgnoreCase);
                Regex regexVariables = new Regex(@"\[[A-Z].*?\]");
                Regex regexVariablesNeg = new Regex(@"\[![A-Z].*?\]"); // they have an exclaimation at the start

                // counts of key words
                int ifCount = regexIfCondition.Matches(docText).Count;
                tboxConditionals.Text = ifCount.ToString();
                int loopCount = regexForEach.Matches(docText).Count;
                tboxLoops.Text = loopCount.ToString();
                int includesCount = regexIncludes.Matches(docText).Count;
                tboxIncludes.Text = includesCount.ToString();
                int variablesCount = regexVariables.Matches(docText).Count;
                int variablesCountNeg = regexVariablesNeg.Matches(docText).Count;
                int totalVariables = variablesCount + variablesCountNeg;
                tboxVariables.Text = totalVariables.ToString();

                // If statements
                //string newWords = Regex.Replace(docText, "\\[&If.[A-Z]*[0-9]*=.*?\\]", htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);
                //newWords = Regex.Replace(newWords, "\\[&Else.?\\]", htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);
                //newWords = Regex.Replace(newWords, "&EndIf", htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);

                // replace conditional statements
                string newWords = regexIfCondition.Replace(docText, htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);
                newWords = regexElseCondition.Replace(newWords, htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);
                newWords = regexEndIfCondition.Replace(newWords, htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);


                // Loops
                newWords = regexForEach.Replace(newWords, htmlTags.tags["div-loop"] + "$&" + htmlTags.tags["div-close"]);
                newWords = regexEndFor.Replace(newWords, htmlTags.tags["div-loop"] + "$&" + htmlTags.tags["div-close"]);

                // Includes
                newWords = regexIncludes.Replace(newWords, htmlTags.tags["div-include"] + "$&" + htmlTags.tags["div-close"]);

                // Update dgIncludes data grid
                Globals.includesList.Clear();
                // Get a collection of matches.
                MatchCollection matches = regexIncludes.Matches(newWords);

                // Use foreach-loop.
                foreach (Match match in matches)
                {
                    Globals.includesList.Add(new mapIncludes() { includeValue = match.Value, includeName = match.Value.Replace("[&amp;Include ", "").Replace("]","") });
                }

                dgIncludes.ItemsSource = Globals.includesList;
                CollectionViewSource.GetDefaultView(dgIncludes.ItemsSource).Refresh();
                dgIncludes.UpdateLayout();
                dgIncludes.Refresh();

                // Variables
                newWords = regexVariables.Replace(newWords, htmlTags.tags["div-field"] + "$&" + htmlTags.tags["div-close"]);
                newWords = regexVariablesNeg.Replace(newWords, htmlTags.tags["div-field"] + "$&" + htmlTags.tags["div-close"]);

                // wrap partial html tags in a htmldocument
                newWords = htmlTags.tags["html-open"] +
                           htmlTags.tags["head-open"] +
                           htmlTags.tags["UTF-8"] +
                           Globals.style +
                           htmlTags.tags["title-open"] + Globals.selectedDocument + htmlTags.tags["title-close"] +
                           htmlTags.tags["head-close"] +
                           htmlTags.tags["body-open"] +
                                newWords +
                           htmlTags.tags["body-close"] +
                           htmlTags.tags["html-close"];

                /*tboxConditionals.Text = ifCount.ToString();
                tboxLoops.Text = loopCount.ToString();
                tboxIncludes.Text = includesCount.ToString();
                tboxVariables.Text = Convert.ToString(variablesCount + variablesCountNeg);
                */

                return newWords;

            } // usingWordProcessingDocument , this closes the doc too
        }


        public void ReplaceMergeFields (string convertedFile)
        {
            // Method open converted docx document in interregoate using open xml to replace merge fields
            // and bring in includes too (converting and then including)

            //https://docs.microsoft.com/en-us/office/open-xml/how-to-search-and-replace-text-in-a-document-part

            int fieldsReplaced = 0;

            using (WordprocessingDocument wordDoc =
            WordprocessingDocument.Open(convertedFile, true))
            {

                // Assign a reference to the existing document body.
                //Body body = wordDoc.MainDocumentPart.Document.Body;
                // create merge field like «Estate_Details_Salutation_ESD01»
                //string wordMergeFieldTxt = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", item.mergeField);
                //SimpleField simpleField1 = new SimpleField() { Instruction = wordMergeFieldTxt };

                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                // As a general rule all single brackets should be replaced with double brackets

                docText = docText.Replace("[", "[[");
                docText = docText.Replace("]", "]]");

                // replace conditional statements
                docText = docText.Replace("&amp;EndIf", "*ENDIF*");
                docText = docText.Replace("&amp;If", "*IF");
                docText = docText.Replace("&amp;Else", "*ELSE*");

                // inspect all mapped fields and replace with converted values by looping through Globals.mergeFieldMapping
                var DistinctItems = Globals.mergeFieldMapping.GroupBy(x => x.solcaseField).Select(y => y.First());

                foreach (var item in DistinctItems)
                {
                    if(item.mergeField != "unmapped")
                    {
                        string solcaseField = @"["+item.solcaseField+@"?]"; // find variable value including square brackets
                        string mergedField = item.mergeField; // add double brackets later?



                        Regex regexText = new Regex(item.solcaseField);
                        docText = regexText.Replace(docText, item.mergeField);

                        fieldsReplaced += 1;
                     
                        //solcaseField = @"[!" + item.solcaseField + @"?]"; // repeat for ! exclamation variables eg [!EXE01]
                        //regexText = new Regex(item.solcaseField);
                        //docText = regexText.Replace(docText, item.mergeField);

                    }
                }

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }

            label1.Content = fieldsReplaced.ToString() + " Fields replaced.";
        }


        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            //label1.Content = "Converting to docx.";
            //string convertedFile = ConvertDocToDocx();
        
            // copy the file to the Convert directory
            string fileName = Globals.selectedDocument;
            string sourcePath = Globals.directoryInfo.FullName;
            string targetPath = System.IO.Path.Combine(sourcePath, "Convert");

            // Use Path class to manipulate file and directory paths.
            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName);

            // To copy a folder's contents to a new location:
            // Create a new target folder.
            // If the directory already exists, this method does not create a new directory.
            System.IO.Directory.CreateDirectory(targetPath);

            // To copy a file to another location and
            // overwrite the destination file if it already exists.
            System.IO.File.Copy(sourceFile, destFile, true);

            label1.Content = "Replacing Merged Fields.";
            ReplaceMergeFields(destFile);
        }
    }

    public static class ExtensionMethods
    {
        private static Action EmptyDelegate = delegate () { };

        public static void Refresh(this UIElement uiElement)
        {
            uiElement.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate);
        }
    }

    public static class Globals
    {
        public static DirectoryInfo directoryInfo { get; set; }

        public static Document activeDoc { get; set; }

        public static string selectedDocument { get; set; }

        public static List<mapMergeField> mergeFieldMapping = new List<mapMergeField>();

        public static List<mapIncludes> includesList = new List<mapIncludes>();

        public static string style = @"<style>
            body{
            font-family: verdana;
            }
            /*Using CSS class for div*/
            .field
            {
            background-color: powderblue;
            }
            .conditional
            {
            background-color: lightgreen;
            }
            .loop
            {
            background-color: orange;
            }
            .include
            {
            background-color: cyan;
            }
            </style>";

        static Globals()
        {

        }
    }

    static class htmlTags
    {
        public static Dictionary<string, string> tags = new Dictionary<string, string>()
        {
            {"break", "<BR />"},
            {"paragraph-open", "<P>"},
            {"paragraph-close", "</P>"},
            {"html-open", "<HTML>"},
            {"html-close", "</HTML>"},
            {"head-open", "<HEAD>"},
            {"head-close", "</HEAD>"},
            {"title-open", "<TITLE>"},
            {"title-close", "</TITLE>"},
            {"body-open", "<BODY>"},
            {"body-close", "</BODY>"},
            {"UTF-8","<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />" },
            {"div-field", "<div class='field'>"},
            {"div-loop", "<div class='loop'>"},
            {"div-include", "<div class='include'>"},
            {"div-conditional", "<div class='conditional'>"},
            {"div-close","</div>"}
        };
    }

    public class mapMergeField
    {
        public string solcaseField { get; set; }
        public string mergeField { get; set; }
    }

    public class mapIncludes
    {
        public string includeValue { get; set; }

        public string includeName { get; set; }
    }

    public static class renderHTML
    {
        public static async Task<string> SearchAndHighlight(string document)
        {

            string newWords = null;


            using (WordprocessingDocument wordDoc =
            WordprocessingDocument.Open(document, true))
            {

                // Assign a reference to the existing document body.
                //Body body = wordDoc.MainDocumentPart.Document.Body;
                // create merge field like «Estate_Details_Salutation_ESD01»
                //string wordMergeFieldTxt = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", item.mergeField);
                //SimpleField simpleField1 = new SimpleField() { Instruction = wordMergeFieldTxt };


                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                //Get all words
                string allWords = docText;

                Regex regexIfCondition = new Regex(@"\[&(?i)If(?-i).[A-Z]*[0-9]*[=,<>].*?\]");
                Regex regexElseCondition = new Regex(@"\[&(?i)Else(?-i).?\]");
                Regex regexEndIfCondition = new Regex(@"\[&(?i)EndIf(?-i).*?\]");
                Regex regexForEach = new Regex(@"\[&(?i)FOREACH(?-i).*?\]");
                Regex regexIncludes = new Regex(@"\[&(?i)Include(?-i).*?\]");
                Regex regexVariables = new Regex(@"\[[A-Z].*?\]");
                Regex regexVariablesNeg = new Regex(@"\[![A-Z].*?\]"); // they have an exclaimation at the start

                // counts of key words
                int ifCount = regexIfCondition.Matches(allWords).Count;
                int loopCount = regexForEach.Matches(allWords).Count;
                int includesCount = regexIncludes.Matches(allWords).Count;
                int variablesCount = regexVariables.Matches(allWords).Count;
                int variablesCountNeg = regexVariablesNeg.Matches(allWords).Count;

                // If statements
                newWords = Regex.Replace(allWords, @"\[&(?i)If(?-i).[A-Z]*[0-9]*=.*?\]", htmlTags.tags["break"] + "$&" + htmlTags.tags["break"]);
                newWords = Regex.Replace(newWords, @"\[&(?i)Else(?-i).?\]", htmlTags.tags["break"] + "$&" + htmlTags.tags["break"]);
                newWords = Regex.Replace(newWords, @"\[&(?i)EndIf(?-i).*?\]", htmlTags.tags["break"] + "$&" + htmlTags.tags["break"]);

                // Loops
                newWords = Regex.Replace(newWords, @"\[&(?i)FOREACH(?-i).*?\]", htmlTags.tags["break"] + "$&" + htmlTags.tags["break"]);
                newWords = Regex.Replace(newWords, @"\[&(?i)ENDFOR(?-i).*?\]", htmlTags.tags["break"] + "$&" + htmlTags.tags["break"]);

                // Includes
                newWords = Regex.Replace(newWords, @"\[&(?i)Include(?-i).*?\]", htmlTags.tags["break"] + "$&" + htmlTags.tags["break"]);

                // Variables
                newWords = Regex.Replace(newWords, @"\[[A-Z].*?\]", htmlTags.tags["div-field"] + "$&" + htmlTags.tags["div-close"]);
                newWords = Regex.Replace(newWords, @"\[![A-Z].*?\]", htmlTags.tags["div-field"] + "$&" + htmlTags.tags["div-close"]);

                // if statements
                newWords = Regex.Replace(newWords, @"\[&(?i)If(?-i).[A-Z]*[0-9]*[=,<>].*?\]", htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);
                newWords = Regex.Replace(newWords, @"\[&(?i)Else(?-i).?\]", htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);
                newWords = Regex.Replace(newWords, @"\[&(?i)EndIf(?-i).*?\]", htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);

                newWords = htmlTags.tags["html-open"] +
                           htmlTags.tags["head-open"] +
                           Globals.style +
                           htmlTags.tags["title-open"] + Globals.selectedDocument + htmlTags.tags["title-close"] +
                           htmlTags.tags["head-close"] +
                           htmlTags.tags["body-open"] +
                                newWords +
                           htmlTags.tags["body-close"] +
                           htmlTags.tags["html-close"];

                /*tboxConditionals.Text = ifCount.ToString();
                tboxLoops.Text = loopCount.ToString();
                tboxIncludes.Text = includesCount.ToString();
                tboxVariables.Text = Convert.ToString(variablesCount + variablesCountNeg);
                */

            }

            return await Task.FromResult(newWords);

        }

        public static string SearchAndHighlightXMLDoc(string document)
        {
            // this method uses Open XML to read the docx version of the template

            //Microsoft.Office.Interop.Word.Application Word97 = new Microsoft.Office.Interop.Word.Application();
            //Word97.WordBasic.DisableAutoMacros();

            using (WordprocessingDocument wordDocx =
                WordprocessingDocument.Open(document, true))
            {

                // Assign a reference to the existing document body.
                //Body body = wordDoc.MainDocumentPart.Document.Body;
                // create merge field like «Estate_Details_Salutation_ESD01»
                //string wordMergeFieldTxt = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", item.mergeField);
                //SimpleField simpleField1 = new SimpleField() { Instruction = wordMergeFieldTxt };

                string docText = null;

                    using (StreamReader sr = new StreamReader(wordDocx.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd();
                    }

                Regex regexIfCondition = new Regex(@"\[&amp;If.[A-Z]*[0-9]*[=,<>].*?\]", RegexOptions.IgnoreCase);
                Regex regexElseCondition = new Regex(@"\[&amp;Else.?\]", RegexOptions.IgnoreCase);
                Regex regexEndIfCondition = new Regex(@"\[&amp;EndIf.*?\]", RegexOptions.IgnoreCase);
                Regex regexForEach = new Regex(@"\[&amp;FOREACH.*?\]", RegexOptions.IgnoreCase);
                Regex regexEndFor = new Regex(@"\[&amp;ENDFOR.*?\]", RegexOptions.IgnoreCase);
                Regex regexIncludes = new Regex(@"\[&amp;Include.*?\]", RegexOptions.IgnoreCase);
                Regex regexVariables = new Regex(@"\[[A-Z].*?\]");
                Regex regexVariablesNeg = new Regex(@"\[![A-Z].*?\]"); // they have an exclaimation at the start

                // counts of key words
                int ifCount = regexIfCondition.Matches(docText).Count;
                
                int loopCount = regexForEach.Matches(docText).Count;
                int includesCount = regexIncludes.Matches(docText).Count;
                int variablesCount = regexVariables.Matches(docText).Count;
                int variablesCountNeg = regexVariablesNeg.Matches(docText).Count;

                // If statements
                //string newWords = Regex.Replace(docText, "\\[&If.[A-Z]*[0-9]*=.*?\\]", htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);
                //newWords = Regex.Replace(newWords, "\\[&Else.?\\]", htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);
                //newWords = Regex.Replace(newWords, "&EndIf", htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);

                // replace conditional statements
                string newWords = regexIfCondition.Replace(docText, htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);
                newWords = regexElseCondition.Replace(newWords, htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);
                newWords = regexEndIfCondition.Replace(newWords, htmlTags.tags["div-conditional"] + "$&" + htmlTags.tags["div-close"]);


                // Loops
                newWords = regexForEach.Replace(newWords, htmlTags.tags["div-loop"] + "$&" + htmlTags.tags["div-close"]);
                newWords = regexEndFor.Replace (newWords, htmlTags.tags["div-loop"] + "$&" + htmlTags.tags["div-close"]);

                // Includes
                newWords =  regexIncludes.Replace(newWords, htmlTags.tags["div-include"] + "$&" + htmlTags.tags["div-close"]);

                // Variables
                newWords = regexVariables.Replace(newWords, htmlTags.tags["div-field"] + "$&" + htmlTags.tags["div-close"]);
                newWords = regexVariablesNeg.Replace(newWords, htmlTags.tags["div-field"] + "$&" + htmlTags.tags["div-close"]);

                // wrap partial html tags in a htmldocument
                newWords = htmlTags.tags["html-open"] +
                           htmlTags.tags["head-open"] +
                           htmlTags.tags["UTF-8"] +
                           Globals.style +
                           htmlTags.tags["title-open"] + Globals.selectedDocument + htmlTags.tags["title-close"] +
                           htmlTags.tags["head-close"] +
                           htmlTags.tags["body-open"] +
                                newWords +
                           htmlTags.tags["body-close"] +
                           htmlTags.tags["html-close"];

                /*tboxConditionals.Text = ifCount.ToString();
                tboxLoops.Text = loopCount.ToString();
                tboxIncludes.Text = includesCount.ToString();
                tboxVariables.Text = Convert.ToString(variablesCount + variablesCountNeg);
                */

                return newWords;

            } // usingWordProcessingDocument , this closes the doc too
        }
    }

}
