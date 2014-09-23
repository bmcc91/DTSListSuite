using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;

namespace DTSListSuite
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static string magicWebPath;
        public static string inputPath;
        public static string outputPath;
        public static string dtsListFile;
        public static string mDtsListFile;
        public static string ccdPgrmFile;
        public static string muPgrmFile;
        public static string changeTFile;
        public static string changeLFile;
        public static string outputFile;
        public static string retireFile;

        public App()
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
                        magicWebPath = childNode.Attributes["value"].Value;
                        break;
                    case "InputPath":
                        inputPath = childNode.Attributes["value"].Value;
                        break;
                    case "OutputPath":
                        outputPath = childNode.Attributes["value"].Value;
                        break;
                    case "DtsList":
                        dtsListFile = childNode.Attributes["value"].Value;
                        break;
                    case "MasterDtsList":
                        mDtsListFile = childNode.Attributes["value"].Value;
                        break;
                    case "CcdPgrms":
                        ccdPgrmFile = childNode.Attributes["value"].Value;
                        break;
                    case "MuPgrms":
                        muPgrmFile = childNode.Attributes["value"].Value;
                        break;
                    case "ChangesT":
                        changeTFile = childNode.Attributes["value"].Value;
                        break;
                    case "ChangesL":
                        changeLFile = childNode.Attributes["value"].Value;
                        break;
                    case "Output":
                        outputFile = childNode.Attributes["value"].Value;
                        break;
                    case "Retire":
                        retireFile = childNode.Attributes["value"].Value;
                        break;
                    default:
                        break;
                }
            }
            dtsListFile = inputPath + dtsListFile;
            mDtsListFile = inputPath + mDtsListFile;
            ccdPgrmFile = inputPath + ccdPgrmFile;
            muPgrmFile = inputPath + muPgrmFile;
            changeTFile = inputPath + changeTFile;
            changeLFile = inputPath + changeLFile;
            outputFile = outputPath + outputFile;
            retireFile = outputPath + retireFile;
        }
    }
}
