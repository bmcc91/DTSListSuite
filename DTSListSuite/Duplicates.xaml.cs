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
            string lastDTS = "";
            dupeListBox.Items.Clear();
            string[] dtsList = File.ReadAllLines(DTSListSuite.App.dtsListFile);
            foreach (string line in dtsList)
            {
                if (line == lastDTS)
                {
                    dupeListBox.Items.Insert(0, line + "\tDuplicate!");
                    flag = true;
                }
                lastDTS = line;
            }
            if (!flag)
                dupeListBox.Items.Insert(0, "No Duplicates!");
            continueButton.IsEnabled = true;
        }
    }
}