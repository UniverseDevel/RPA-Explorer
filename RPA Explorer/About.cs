using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace RPA_Explorer
{
    public partial class About : Form
    {
        private static string[] translatorsList =
        {
            "-"
        };
        private static string[] contributorsList =
        {
            "-"
        };
        
        private readonly string appVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion;
        private string appCreator = "Martin Suchy";
        private string appRepository = "https://github.com/UniverseDevel/RPA-Explorer";
        private string appTranslators = String.Join(", ", translatorsList);
        private string appContributors = String.Join(", ", contributorsList);
        
        public About()
        {
            InitializeComponent();

            LoadTexts();
        }

        private void LoadTexts()
        {
            Text = RpaExplorer.GetText("About");
            richTextBox1.Text = string.Format(RpaExplorer.GetText("About_text"), appVersion, appCreator, appRepository, appTranslators, appContributors);
        }

        private void richTextBox1_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            Process.Start(e.LinkText);
        }

        private void richTextBox1_Enter(object sender, EventArgs e)
        {
            ActiveControl = null;
        }
    }
}