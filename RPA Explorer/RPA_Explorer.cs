using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using NeoSmart.PrettySize;

namespace RPA_Explorer
{
    public partial class RpaExplorer : Form
    {
        private RpaParser rpaParser;
        private Thread opration;
        private bool operationEnabled = true;
        private SortedDictionary<string, RpaParser.ArchiveIndex> fileList = new ();
        
        public RpaExplorer()
        {
            InitializeComponent();
        }

        private void exportFiles(List<string> exportFilesList, FolderBrowserDialog folderBrowserDialog)
        {
            int counter = 0;
            int jobSize = exportFilesList.Count;
            foreach (string file in exportFilesList)
            {
                counter++;
                int pctProcessed = (int) Math.Ceiling((double) counter / jobSize * 100);
                label1.PerformSafely(() => label1.Text = counter + " / " + jobSize);
                progressBar1.PerformSafely(() => progressBar1.Value = pctProcessed);
                
                statusBar1.PerformSafely(() => statusBar1.Text = "Exporting file: " + file);
                        
                rpaParser.Extract(file, folderBrowserDialog.SelectedPath);

                if (!operationEnabled)
                {
                    break;
                }
            }
                    
            button1.PerformSafely(() => button1.Enabled = true);
            button2.PerformSafely(() => button2.Enabled = true);
            button3.PerformSafely(() => button3.Enabled = true);
            button4.PerformSafely(() => button4.Enabled = true);
            progressBar1.PerformSafely(() => progressBar1.Value = 0);
            label1.PerformSafely(() => label1.Text = "");
            statusBar1.PerformSafely(() => statusBar1.Text = "Ready");
            
            operationEnabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select RenPy Archive";
                openFileDialog.Filter = "RPA/RPI files (*.rpa,*.rpi)|*.rpa;*.rpi)";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (openFileDialog.CheckFileExists)
                    {
                        statusBar1.Text = "Loading file: " + openFileDialog.FileName;
                        
                        rpaParser = new RpaParser(openFileDialog.FileName);

                        fileList = rpaParser.GetFileList();

                        listView1.Items.Clear();
                        foreach (KeyValuePair<string, RpaParser.ArchiveIndex> kvp in fileList)
                        {
                            ListViewItem item = new ListViewItem(kvp.Key);
                            item.Name = kvp.Key;
                            item.SubItems.Add(PrettySize.Format(kvp.Value.length));
                            listView1.Items.Add(item);
                        }

                        textBox1.Text = "File location: " + rpaParser.GetArchiveInfo().FullName + Environment.NewLine +
                                        "File size: " + PrettySize.Format(rpaParser.GetArchiveInfo().Length) + Environment.NewLine +
                                        "Object count: " + rpaParser.GetFileList().Count + Environment.NewLine;

                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                        textBox2.Enabled = true;

                        statusBar1.Text = "Ready";
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
                List<string> exportFilesList = new List<string>();
                foreach (ListViewItem item in listView1.Items)
                {
                    if (item.Checked)
                    {
                        exportFilesList.Add(item.Text);
                    }
                }
                
                if (exportFilesList.Count == 0)
                {
                    return;
                }
                
                folderBrowserDialog.SelectedPath = rpaParser.GetArchiveInfo().DirectoryName;
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    button1.Enabled = false;
                    button2.Enabled = false;
                    button3.Enabled = false;
                    button4.Enabled = false;
                    button5.Enabled = true;
                    progressBar1.Value = 0;
                    label1.PerformSafely(() => label1.Text = "");
                    
                    opration = new Thread(() => exportFiles(exportFilesList, folderBrowserDialog));
                    opration.Start();
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            operationEnabled = false;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (fileList.Count == 0)
            {
                return;
            }

            statusBar1.Text = "Filtering...";
            
            string filter = textBox2.Text;
            if (filter.Length < 2)
            {
                filter = String.Empty;
            }
            
            Regex rgx = new Regex("");
            
            try
            {
                rgx = new Regex(textBox2.Text);
            }
            catch {
                filter = String.Empty;
            }

            foreach (KeyValuePair<string, RpaParser.ArchiveIndex> kvp in fileList)
            {
                if (rgx.IsMatch(kvp.Key) || filter == String.Empty)
                {
                    if (!listView1.Items.ContainsKey(kvp.Key))
                    {
                        ListViewItem item = new ListViewItem(kvp.Key);
                        item.Name = kvp.Key;
                        item.SubItems.Add(PrettySize.Format(kvp.Value.length));
                        listView1.Items.Add(item);
                    }
                }
                else
                {
                    listView1.Items.RemoveByKey(kvp.Key);
                }
            }

            statusBar1.Text = "Ready";
        }
    }
    
    public static class CrossThreadExtensions
    {
        public static void PerformSafely(this Control target, Action action)
        {
            if (target.InvokeRequired)
            {
                target.Invoke(action);
            }
            else
            {
                action();
            }
        }

        public static void PerformSafely<T1>(this Control target, Action<T1> action,T1 parameter)
        {
            if (target.InvokeRequired)
            {
                target.Invoke(action, parameter);
            }
            else
            {
                action(parameter);
            }
        }

        public static void PerformSafely<T1,T2>(this Control target, Action<T1,T2> action, T1 p1,T2 p2)
        {
            if (target.InvokeRequired)
            {
                target.Invoke(action, p1,p2);
            }
            else
            {
                action(p1,p2);
            }
        }
    }
}