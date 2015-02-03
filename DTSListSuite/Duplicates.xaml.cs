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
using System.IO;
using Microsoft.VisualBasic.FileIO;

namespace DTSListSuite
{
    /// <summary>
    /// Interaction logic for Duplicates.xaml
    /// </summary>
    public partial class Duplicates : Page
    {
        public Duplicates()
        {
            InitializeComponent();
        }

        private void Continue_Click(object sender, RoutedEventArgs e)
        {
            continueButton.IsEnabled = false;
            bool flag = false;
            string[] lastDTS = {""};
            dupeListBox.Items.Clear();
            string[] dtsList = File.ReadAllLines(DTSListSuite.App.mDtsListFile);
            foreach (string line in dtsList)
            {
                if (line.Substring(0, 3) != "C/R")
                {
                    TextFieldParser parser = new TextFieldParser(new StringReader(line));
                    parser.HasFieldsEnclosedInQuotes = true;
                    parser.SetDelimiters(",");
                    string[] ArLn = { "" };
                    while (!parser.EndOfData)
                        ArLn = parser.ReadFields();
                    if (lastDTS.Length > 1)
                    {
                        if (ArLn[1] == lastDTS[1] && ArLn[2] == lastDTS[2] && ArLn[3] == lastDTS[3])
                        {
                            dupeListBox.Items.Insert(0, ArLn[1] + " " + ArLn[2] + " " + ArLn[3] + "\tDuplicate!");
                            flag = true;
                        }
                    }
                    lastDTS = ArLn;
                }
            }
            if (!flag)
                dupeListBox.Items.Insert(0, "No Duplicates!");
            continueButton.IsEnabled = true;
        }
    }
}