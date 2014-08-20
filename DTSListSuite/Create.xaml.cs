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
using System.Diagnostics;
using System.Net;
using System.IO;
using Microsoft.VisualBasic.FileIO;

namespace DTSListSuite
{
    /// <summary>
    /// Interaction logic for Create.xaml
    /// </summary>
    public partial class Create : Page
    {
        public Create()
        {
            InitializeComponent();
        }
        private async void Compile_Click(object sender, RoutedEventArgs e)
        {
            /* Start compiling, will never be called multiple times at once */
            compileButton.IsEnabled = false;
            compileButton.Content = "Compiling...";
            bool check = await checkIntranet();
            if (check)
            {
                /* Check settings */
                string[] dtsList = readData(DTSListSuite.App.dtsListFile);
                string[] csvList = readData(DTSListSuite.App.mDtsListFile);
                string[] ccdPgrms = readData(DTSListSuite.App.ccdPgrmFile);
                string[] mu2Pgrms = readData(DTSListSuite.App.muPgrmFile);
                string[] changesT = readData(DTSListSuite.App.changeTFile);
                string[] changesL = readData(DTSListSuite.App.changeLFile);
                compileBar.Maximum = dtsList.Length;
                /* Compile and output */
                List<string[]> sheetData = await compileData(dtsList, ccdPgrms, mu2Pgrms, changesT, changesL);
                outputData(sheetData, csvList);
            }
            else
                MessageBox.Show("Failure to connect to Magic Web database", "Failure to Connect",
                    MessageBoxButton.OK, MessageBoxImage.Hand);
            compileButton.IsEnabled = true;
            compileButton.Content = "Compile";
        }

        private async Task<bool> checkIntranet()
        {
            /* Sorta slow but whatever */
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(DTSListSuite.App.magicWebPath);
            request.AllowAutoRedirect = false;
            request.Method = "HEAD";
            try
            {
                WebResponse response = await request.GetResponseAsync();
            }
            catch (WebException)
            {
                return (false);
            }
            return (true);
        }

        private string[] readData(string input)
        {
            string[] dtss = File.ReadAllLines(input);
            return (dtss);
        }

        private async Task<List<string[]>> compileData(string[] dtsList, string[] ccdPgrms, string[] mu2Pgrms,
            string[] changesT, string[] changesL)
        {
            List<string[]> returnList = new List<string[]> { };
            int i = 0;
            foreach (string dts in dtsList)
            {
                i++;
                WebClient client = new WebClient();
                client.DownloadStringCompleted += (s, e) => 
                    {
                        dtsSlot temp = new dtsSlot(e.Result, ccdPgrms, mu2Pgrms, changesT, changesL);
                        compileBar.Value = i;
                        /* Add to output list */
                        returnList.Add(temp.results);
                    };
                string[] dtsArray = dts.Split('\t');
                var url = string.Format(DTSListSuite.App.magicWebPath + "dts/REQUESTS/{0}/{1}/{2}.htm",
                    dtsArray[0], dtsArray[1], dtsArray[2]);
                await client.DownloadStringTaskAsync(new Uri(url));
            }
            return (returnList);
        }

        private void outputData(List<string[]> sheetData, string[] dtsList)
        {
            int capacity = sheetData.Count;
            string[] printList = new string[capacity];
            int i = 1;
            foreach (string[] item in sheetData)
            {
                /* Start Temp Replacements */
                TextFieldParser parser = new TextFieldParser(new StringReader(dtsList[i]));
                parser.HasFieldsEnclosedInQuotes = true;
                parser.SetDelimiters(",");
                string[] tempDTS = {""};
                while (!parser.EndOfData)
                    tempDTS = parser.ReadFields();
                if (item[4] == "Submitted")
                    item[9] = tempDTS[10];
                if (item[0] == "FOC")
                {
                    item[14] = tempDTS[15];
                    item[15] = tempDTS[16];
                }
                if ((item[14] == "Y") || (item[15] == "Y"))
                    item[13] = "";
                item[16] = tempDTS[17];
                if (tempDTS[18] == "Y")
                    item[17] = tempDTS[18];
                if (tempDTS[19] == "Y")
                    item[18] = tempDTS[19];
                /* End Temp Replacements */
                string tempItem = string.Join("\t", item);
                printList[i - 1] = tempItem;
                i++;
            }
            File.WriteAllLines(DTSListSuite.App.outputFile, printList);
        }
    }

    public class dtsSlot
    {
        public string[] results = new string[19];

        public dtsSlot(string dtsHTML, string[] ccd, string[] mu2, string[] changesT, string[] changesL)
        {
            DateTime dateTime = DateTime.Today;
            string[] basics = setupBasicInfo(dtsHTML);
            results[0] = basics[0];
            results[1] = basics[1];
            results[2] = basics[2];
            results[3] = description(dtsHTML);
            results[4] = status(dtsHTML);
            results[5] = priority(dtsHTML);
            results[6] = associated(dtsHTML);
            results[7] = release(dtsHTML);
            results[8] = ppack(dtsHTML);
            results[9] = program(dtsHTML);
            results[10] = dateTime.ToString("MM/dd/yyyy");
            results[11] = change(dtsHTML);
            results[12] = duedate(dtsHTML);
            results[13] = pushShip(dtsHTML, "6.07C", results[8]);
            results[14] = shipRing(results[11], changesT);
            results[15] = shipRing(results[11], changesL);
            results[16] = "";
            results[17] = compareList(ccd, results[9], dtsHTML);
            if (results[17] != "Y")
                results[18] = compareList(mu2, results[9], dtsHTML);
        }

        private string parseTag(string tag, string dtsHTML)
        {
            /* Many properties can be parsed by the corresponding HTML header value
             * For those that can, no other methods are needed */
            int i = dtsHTML.IndexOf(tag);
            if (i != -1)
            {
                dtsHTML = dtsHTML.Substring(i);
                i = dtsHTML.IndexOf("content") + 9;
                dtsHTML = dtsHTML.Substring(i);
                int j = dtsHTML.IndexOf("\">");
                dtsHTML = dtsHTML.Substring(0,j);
                return (dtsHTML);
            }
            else
                return ("");
        }

        private string[] setupBasicInfo(string dtsHTML)
        {
            string product = parseTag("prodln", dtsHTML);
            string application = parseTag("appl", dtsHTML);
            string dtsNum = parseTag("DTSnumber", dtsHTML);
            /* Added quotations for automatic hyperlinking in google spreadsheets */
            dtsNum = "=hyperlink(\"\"\"\"http://magicweb/dts/REQUESTS/" + product +
                     "/" + application + "/" + dtsNum + ".htm\"\"\"\"," + dtsNum + ")";
            return(new string[] {product, application, dtsNum});
        }

        private string description(string dtsHTML)
        {
            string content = parseTag("description", dtsHTML);
            if (content != "")
            {
                /* Replace HTML quotes with ASCII */
                content = content.Substring(0, content.Length - 4);
                int i = content.IndexOf("&quot;");
                while (i != -1)
                {
                    string tempA = content.Substring(0, i) + "\"";
                    string tempB = content.Substring(i + 6);
                    content = tempA + tempB;
                    i = content.IndexOf("&quot;");
                }
                return (content);
            }
            else
                return ("");
        }
        
        private string status(string dtsHTML)
        {
            /* Can probably clean this up later */
            bool flag = false;
            string pri = parseTag("DTSreleasesVIEW", dtsHTML);
            int i = pri.IndexOf("6.07");
            if (i == -1)
            {
                i = pri.IndexOf("6.1");
                if (i == -1)
                    return ("");
                flag = true;
                int n;
                bool isNumeric = int.TryParse(pri.Substring(i + 3, 1), out n);
                if (isNumeric)
                    i++;
                pri = pri.Substring(i + 3, 1);
            }
            else
                pri = pri.Substring(i + 4, 1);
            switch (pri)
            {
                case "S":
                    pri = "Submitted";
                    break;
                case "D":
                    pri = "Draft";
                    break;
                case "T":
                    pri = "Testing";
                    break;
                case "Q":
                    pri = "Queued";
                    break;
                case "U":
                    pri = "United Tested";
                    break;
                case "P":
                    pri = "Production";
                    break;
                case "C":
                    pri = "Completed";
                    break;
                case "R":
                    pri = "Rejected";
                    break;
                case "r":
                    pri = "Reclass";
                    break;
                case "H":
                    pri = "Holding";
                    break;
                default:
                    break;
            }
            if (flag)
                pri = "6.1-" + pri;
            return (pri);
        }

        private string priority(string dtsHTML)
        {
            return (parseTag("DTSpriority", dtsHTML));
        }
        
        private string associated(string dtsHTML)
        {
            int i = dtsHTML.IndexOf("<th>Req'd/Link</th>");
            if (i == -1)
                return ("");
            dtsHTML = dtsHTML.Substring(i);
            i = dtsHTML.IndexOf("data\">") + 6;
            dtsHTML = dtsHTML.Substring(i,1);
            if (dtsHTML == "Y")
                return (dtsHTML);
            return ("");
        }

        private string release(string dtsHTML)
        {
            string content = parseTag("DTSreleasesVIEW", dtsHTML);
            if (content != "")
            {
                int i = content.IndexOf("6.");
                if (i != -1)
                {
                    content = content.Substring(i);
                    return (content);
                }
                else
                    return ("");
            }
            else
                return ("");
        }

        private string ppack(string dtsHTML)
        {
            string content = parseTag("DTSSRnumbers", dtsHTML);
            if (content != "")
            {
                int i = dtsHTML.IndexOf("6.07SR");
                if (i != -1)
                {
                    dtsHTML = dtsHTML.Substring(i + 6, 1);
                    return (dtsHTML);
                }
                else
                    return ("");
            }
            else
                return ("");
        }
        
        private string program(string dtsHTML)
        {
            /* Lots of magic numbers...will fix later */
            /* I hate parsing raw javascript */
            int i = dtsHTML.IndexOf("ToggleAllCompDoc") + 151;
            dtsHTML = dtsHTML.Substring(i,2000);
            i = dtsHTML.IndexOf("mySections[");
            if (i != -1)
            {
                int j = dtsHTML.IndexOf("(i=0") - 4;
                dtsHTML = dtsHTML.Substring(0, j);
                /* At this point, dtsHTML is a list of unformatted programs only */
                string[] delims = new string[] { "\n" };
                string[] dtsList = dtsHTML.Split(delims, StringSplitOptions.RemoveEmptyEntries);
                List<string> returnList = new List<string>();
                foreach (string item in dtsList)
                {
                    if (item.Length > 2)
                        returnList.Add(item);
                }
                dtsHTML = "";
                foreach (string cell in returnList)
                {
                    string tempCell = cell;
                    tempCell = tempCell.Substring(18);
                    tempCell = tempCell.Substring(0, tempCell.Length - 7);
                    int tempCut = tempCell.IndexOf("\"") + 1;
                    tempCell = tempCell.Substring(tempCut);
                    dtsHTML += (tempCell + ", ");
                }
                if (dtsHTML.Length > 0)
                    dtsHTML = dtsHTML.Substring(0, dtsHTML.Length - 2);
                return (dtsHTML);
            }
            else
                return ("");
        }

        private string change(string dtsHTML)
        {
            string appl = parseTag("appl", dtsHTML);
            if (appl != "")
            {
                string content = parseTag("DTSRelCh607", dtsHTML);
                if (content != "")
                {
                    content = content.Substring(2);
                    return (appl + " " + content);
                }
                else
                    return ("");
            }
            else
                return ("");
        }

        private string duedate(string dtsHTML)
        {
            string content = parseTag("DTSduedate", dtsHTML);
            if (content != "")
                return (content);
            else
                return ("");
        }

        private string pushShip(string dtsHTML, string release, string ppack)
        {
            int i = dtsHTML.IndexOf(release);
            if ((i != -1) && (ppack == ""))
                return ("Y");
            else
                return ("");
        }

        private string compareList(string[] list, string programs, string dtsHTML)
        {
            string keywords = parseTag("DTSkeywords", dtsHTML);
            foreach (string line in list)
            {
                if (programs.IndexOf(line) != -1)
                    return ("Y");
                if (keywords.IndexOf(line) != -1)
                    return ("Y");
            }
            return ("");
        }

        private string shipRing(string change, string[] changeList)
        {
            if (change == "")
                return ("");
            string appl = change.Substring(0, 3);
            change = change.Substring(4);
            foreach(string line in changeList)
            {
                string tempAppl = line.Substring(0, 3);
                if (appl == tempAppl)
                {
                    string[] tempLine = line.Split(new String[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                    if (Convert.ToInt32(change) <= Convert.ToInt32(tempLine[1]))
                        return ("Y");
                    foreach (string item in tempLine)
                    {
                        if (change == item)
                            return ("Y");
                    }
                }
            }
            return ("");
        }
    }
}
