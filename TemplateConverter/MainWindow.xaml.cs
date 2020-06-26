using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            string myMessage = "";

            if (treeView1.Items.Count >= 0)
            {
                var tree = sender as System.Windows.Controls.TreeView;

                if (tree.SelectedItem is TreeViewItem)
                {
                    // ... Handle a TreeViewItem.
                    var item = tree.SelectedItem as TreeViewItem;
                    myMessage = item.Header.ToString();
                }
                else if (tree.SelectedItem is string)
                {
                    // ... Handle a string.
                    myMessage = tree.SelectedItem.ToString();
                }
            }

            string document = System.IO.Path.Combine(Globals.directoryInfo.FullName, myMessage);
            //MessageBoxResult result = System.Windows.MessageBox.Show(myMessage);

            SearchAndHighlight(document, myMessage);
        }

        private void SearchAndHighlight(string document, string fileName)
        {
            label1.Content = "Initiating MS Word.";

            Microsoft.Office.Interop.Word.Application Word97 = new Microsoft.Office.Interop.Word.Application();
            //Word97.WordBasic.DisableAutoMacros();

            label1.Content = "Opening word document.";

            Document doc = new Document();

            try
            {
                doc = Word97.Documents.Open(document);
            }
            catch (Exception ex)
            {
                MessageBoxResult exception = System.Windows.MessageBox.Show(ex.Message);
            }


            //Get all words
            string allWords = doc.Content.Text;

            // close the document, no need for it open now
            doc.Close();
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
                       htmlTags.tags["title-open"] + fileName + htmlTags.tags["title-close"] +
                       htmlTags.tags["head-close"] +
                       htmlTags.tags["body-open"] +
                            newWords + 
                       htmlTags.tags["body-close"] +
                       htmlTags.tags["html-close"];

            htmlOutput.NavigateToString(newWords);
          
            label1.Content = "";

            /*tboxConditionals.Text = ifCount.ToString();
            tboxLoops.Text = loopCount.ToString();
            tboxIncludes.Text = includesCount.ToString();
            tboxVariables.Text = Convert.ToString(variablesCount + variablesCountNeg);
            */

        }

    }



    public static class Globals
    {
        public static DirectoryInfo directoryInfo { get; set; }

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

}
