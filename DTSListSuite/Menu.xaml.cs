using System;
using System.IO;
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

namespace DTSListSuite
{
    /// <summary>
    /// Interaction logic for Menu.xaml
    /// </summary>
    public partial class Menu : Page
    {
        public Menu()
        {
            InitializeComponent();
        }
        private void Continue_Click(object sender, RoutedEventArgs e)
        {
            string selection;
            bool filesExist = checkFiles();
            if ( (ListBoxItem)peopleListBox.SelectedItem == null)
            {
                selection = "";
            }
            else
            {
                selection = ((ListBoxItem)peopleListBox.SelectedValue).Content.ToString();
            }
            switch(selection)
            {
                /* Logic for each menu choice */
                case "Create List":
                    if (filesExist)
                    {
                        Create createScreen = new Create();
                        this.NavigationService.Navigate(createScreen);
                    }
                    else
                        MessageBox.Show("Missing files necessary for functionality",
                            "Missing Files", MessageBoxButton.OK, MessageBoxImage.Error);
                    break;
                case "Find Duplicates":
                    if (filesExist)
                    {
                        Duplicates dupeScreen = new Duplicates();
                        this.NavigationService.Navigate(dupeScreen);
                    }
                    else
                        MessageBox.Show("Missing files necessary for functionality",
                            "Missing Files", MessageBoxButton.OK, MessageBoxImage.Error);
                    break;
                case "Open DTS List":
                    Process.Start("https://docs.google.com/a/meditech.com/spreadsheets/d/1jriJKvq-BtTI0yGvecFnDqftwE4j5tZR_2lYBI6XQ9g/edit#gid=0");
                    break;
                case "Options":
                    Options optionsScreen = new Options();
                    this.NavigationService.Navigate(optionsScreen);
                    break;
                case "README":
                    Process.Start("README.txt");
                    break;
                default:
                    Debug.WriteLine("Nothing");
                    break;
            }
        }
        private bool checkFiles()
        {
            /* Check if the necessary files exist before running */
            bool ccd = File.Exists(DTSListSuite.App.ccdPgrmFile);
            bool mut = File.Exists(DTSListSuite.App.muPgrmFile);
            bool liv = File.Exists(DTSListSuite.App.changeLFile);
            bool tes = File.Exists(DTSListSuite.App.changeTFile);
            bool mas = File.Exists(DTSListSuite.App.mDtsListFile);
            if (ccd && mut && liv && tes && mas)
                return (true);
            else
                return (false);
        }
    }
}
