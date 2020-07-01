﻿using System;
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
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;
using HtmlAgilityPack;
using System.Data.OleDb;

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
                    //handle error
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

        private async void treeView1_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {

            label1.Content = "Initiating MS Word.";

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

            string htmlText = "";

            try
            {
                htmlText = await renderHTML.SearchAndHighlight(document);
                btnConvert.IsEnabled = true;
            } 
            catch (Exception ex)
            {
                htmlText = "<HTML><BODY><H1> Unable to display document! </H1> <br />" + ex.Message + "</BODY><?HTML>";
            }

            // load lookup
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("C:\\Users\\kieran.caulfield\\OneDrive - Birkett Long LLP\\Documents\\Transformation\\Auto-Match-Screen-Fields.xml");

            htmlOutput.NavigateToString(htmlText);
            label1.Content = Globals.selectedDocument;

            // find all variables identified in our displayed html (where div class is 'field')
            var divFieldsXPath = "//div[contains(@class,'field')]";

            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(htmlText);

            HtmlNodeCollection listOfFields = htmlDoc.DocumentNode.SelectNodes(divFieldsXPath);
            // check there are actually some fields to iterate round

            if (listOfFields != null)
            {
                foreach (var node in listOfFields)
                {
                    // find the Merge Field Mapping
                    Regex fieldName = new Regex(@"[A-Z][A-Z]?[A-Z][0-9]."); // will find TF09 and ABC09

                    Match m = fieldName.Match(node.InnerText, 0);

                    string mappedMergeField = "unmapped";
                    // search xml merged field table with this xpath to get match "//Auto-Match-Screen-Fields/Data-Collection-Field-Name[../Field-Code-Lookup='ESD03']/text()"

                    XmlNode mappedNode = xmlDoc.SelectSingleNode("//Auto-Match-Screen-Fields/Data-Collection-Field-Name[../Field-Code-Lookup='" + m.Value + "']/text()");

                    if (mappedNode != null)
                    {
                        mappedMergeField = mappedNode.InnerText;
                    }

                    Globals.mergeFieldMapping.Add(new mapMergeField() { solcaseField = m.Value, mergeField = mappedMergeField });
                }
            }

            // Bind list to data grid
            dgMap.ItemsSource = Globals.mergeFieldMapping;

        }



        public void ConvertDocToDocx()
        {
            // https://stackoverflow.com/questions/23120431/saveas-ms-word-as-docx-instead-of-doc-c-sharp

            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            string document = System.IO.Path.Combine(Globals.directoryInfo.FullName, Globals.selectedDocument);

            string newFileName = "";

            try
            {
                Globals.activeDoc = word.Documents.Open(document);



                if (Globals.activeDoc.FullName.ToLower().EndsWith(".doc"))
                {
                    if (Globals.activeDoc.FullName.EndsWith(".DOC"))
                    {
                        newFileName = Globals.activeDoc.FullName.Replace(".DOC", ".AS.docx");
                    } else
                    {
                        newFileName = Globals.activeDoc.FullName.Replace(".doc", ".AS.docx");
                    }

                    Globals.activeDoc.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                                     CompatibilityMode: WdCompatibilityMode.wdWord2013);
                }
            }
            catch (Exception ex)
            {
                MessageBoxResult exception = System.Windows.MessageBox.Show(ex.Message);
            }

            label1.Content = Globals.activeDoc.Name;

            word.ActiveDocument.Close();
            word.Quit();

        }

        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            label1.Content = "Converting to docx.";
            ConvertDocToDocx();

        }
    }



    public static class Globals
    {
        public static DirectoryInfo directoryInfo { get; set; }

        public static Document activeDoc { get; set; }

        public static string selectedDocument { get; set; }

        public static List<mapMergeField> mergeFieldMapping = new List<mapMergeField>();

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
            background-color: lightorange;
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

    public static class renderHTML
    {
        public static async Task<string> SearchAndHighlight(string document)
        {

            Microsoft.Office.Interop.Word.Application Word97 = new Microsoft.Office.Interop.Word.Application();
            //Word97.WordBasic.DisableAutoMacros();

            // Document doc = new Document();

            try
            {
                Globals.activeDoc = Word97.Documents.Open(document);
            }
            catch (Exception ex)
            {
                MessageBoxResult exception = System.Windows.MessageBox.Show(ex.Message);
            }

            //Get all words
            string allWords = Globals.activeDoc.Content.Text;

            // close the document, no need for it open now
            Globals.activeDoc.Close();
            Word97.Quit();

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
            string newWords = Regex.Replace(allWords, @"\[&(?i)If(?-i).[A-Z]*[0-9]*=.*?\]", htmlTags.tags["break"] + "$&" + htmlTags.tags["break"]);
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

            return await Task.FromResult(newWords);

        }
    }

}
