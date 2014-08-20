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
using System.Xml;

namespace DTSListSuite
{
    /// <summary>
    /// Interaction logic for Options.xaml
    /// </summary>
    public partial class Options : Page
    {
        public Options()
        {
            InitializeComponent();
            webPathBox.Text = DTSListSuite.App.magicWebPath;
            inputPathBox.Text = DTSListSuite.App.inputPath;
            outputPathBox.Text = DTSListSuite.App.outputPath;
        }

        private void Continue_Click(object sender, RoutedEventArgs e)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(AppDomain.CurrentDomain.BaseDirectory + "..\\..\\App.config");
            XmlNode appSettingsNode = xmlDoc.SelectSingleNode("configuration/appSettings");
            foreach (XmlNode childNode in appSettingsNode)
            {
                string selection = childNode.Attributes["key"].Value;
                switch (selection)
                {
                    case "MagicWebPath":
                        childNode.Attributes["value"].Value = webPathBox.Text;
                        break;
                    case "InputPath":
                        childNode.Attributes["value"].Value = inputPathBox.Text;
                        break;
                    case "OutputPath":
                        childNode.Attributes["value"].Value = outputPathBox.Text;
                        break;
                    default:
                        break;
                }              
            }
            xmlDoc.Save(AppDomain.CurrentDomain.BaseDirectory + "..\\..\\App.config");
            xmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
        }
    }
}
