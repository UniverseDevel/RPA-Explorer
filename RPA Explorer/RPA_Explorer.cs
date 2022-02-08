using System;
using System.Collections.Generic;
using System.Windows.Forms;
using NeoSmart.PrettySize;

namespace RPA_Explorer
{
    public partial class RpaExplorer : Form
    {
        private RpaParser rpaParser;
        
        public RpaExplorer()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select RenPy Archive";
                openFileDialog.Filter = "RPA files (*.rpa)|*.rpa|RPI files (*.rpi)|*.rpi)";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (openFileDialog.CheckFileExists)
                    {
                        rpaParser = new RpaParser(openFileDialog.FileName);

                        SortedDictionary<string, RpaParser.ArchiveIndex> fileList = rpaParser.GetFileList();

                        listView1.Items.Clear();
                        foreach (KeyValuePair<string, RpaParser.ArchiveIndex> kvp in fileList)
                        {
                            ListViewItem item = new ListViewItem(kvp.Key);
                            item.SubItems.Add(PrettySize.Format(kvp.Value.length));
                            listView1.Items.Add(item);
                        }

                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.Items)
            {
                item.Checked = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.Items)
            {
                item.Checked = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.SelectedPath = rpaParser.GetArchiveInfo().DirectoryName;
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (ListViewItem item in listView1.Items)
                    {
                        if (item.Checked)
                        {
                            rpaParser.Extract(item.Text, folderBrowserDialog.SelectedPath);
                        }
                    }
                }
            }
        }
    }
}