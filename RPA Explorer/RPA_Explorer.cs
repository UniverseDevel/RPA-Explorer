using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using LibVLCSharp.Shared;
using NeoSmart.PrettySize;
using RPA_Parser;

namespace RPA_Explorer
{
    public partial class RpaExplorer : Form
    {
        private RpaParser rpaParser;
        private Thread opration;
        private bool operationEnabled = true;
        private SortedDictionary<string, RpaParser.ArchiveIndex> fileList = new ();
        private string[] args;
        private bool switchTabs = false;
        private LibVLC libVlc;
        private MemoryStream memoryStreamVlc;
        private StreamMediaInput streamMediaInputVlc;
        private Media mediaVlc;

        public RpaExplorer()
        {
            InitializeComponent();
            Core.Initialize();
            libVlc = new LibVLC("--input-repeat=2");

            args = Environment.GetCommandLineArgs();
            if (args.Length >= 2)
            {
                loadArchive(args[1]);
            }
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
            button3.PerformSafely(() => button3.Enabled = false);
            progressBar1.PerformSafely(() => progressBar1.Value = 0);
            label1.PerformSafely(() => label1.Text = "");
            statusBar1.PerformSafely(() => statusBar1.Text = "Ready");
            
            operationEnabled = true;
        }

        private void loadArchive(string openFile)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select RenPy Archive";
                openFileDialog.Filter = "RPA/RPI files (*.rpa,*.rpi)|*.rpa;*.rpi)";

                DialogResult dialogResult = DialogResult.None;
                if (openFile == String.Empty)
                {
                    dialogResult = openFileDialog.ShowDialog();
                }
                else
                {
                    if (openFile.EndsWith(".rpa") || openFile.EndsWith(".rpi"))
                    {
                        dialogResult = DialogResult.OK;
                        openFileDialog.FileName = openFile;
                    }
                }

                if (dialogResult == DialogResult.OK)
                {
                    if (openFileDialog.CheckFileExists)
                    {
                        statusBar1.Text = "Loading file: " + openFileDialog.FileName;
                        
                        rpaParser = new RpaParser(openFileDialog.FileName);

                        fileList = rpaParser.GetFileList();

                        treeView1.Nodes.Clear();
                        TreeNode root = new TreeNode();
                        TreeNode node = root;
                        root.Name = "";
                        root.Text = "/";
                        root.BackColor = Color.SandyBrown;
                        treeView1.Nodes.Add(root);
                        string pathBuild = String.Empty;
                        foreach (KeyValuePair<string, RpaParser.ArchiveIndex> kvp in fileList)
                        {
                            node = root;
                            pathBuild = String.Empty;
                            foreach (string pathBits in kvp.Key.Split('/'))
                            {
                                if (pathBuild == String.Empty)
                                {
                                    pathBuild = pathBits;
                                }
                                else
                                {
                                    pathBuild += "/" + pathBits;
                                }

                                if (node.Nodes.ContainsKey(pathBits))
                                {
                                    node = node.Nodes[pathBits];
                                }
                                else
                                {
                                    string sizeInfo = String.Empty;
                                    if (fileList.ContainsKey(pathBuild))
                                    {
                                        sizeInfo = " (" + PrettySize.Format(fileList[pathBuild].length) + ")";
                                    }
                                    else
                                    {
                                        // Loop trough fileList and find all .StartsWith(pathBuild) keys and sum their .length
                                    }
                                    node = node.Nodes.Add(pathBits, pathBits + sizeInfo);
                                    if (!fileList.ContainsKey(pathBuild))
                                    {
                                        node.BackColor = Color.SandyBrown;
                                    }
                                }
                            }
                        }
                        root.Expand();

                        textBox1.Text = "Archive version: " + rpaParser.GetArchiveVersion() + Environment.NewLine +
                                        "File location: " + rpaParser.GetArchiveInfo().FullName + Environment.NewLine +
                                        "File size: " + PrettySize.Format(rpaParser.GetArchiveInfo().Length) + Environment.NewLine +
                                        "Object count: " + fileList.Count + Environment.NewLine;

                        switchTabs = true;
                        tabControl1.SelectedTab = tabPage0;
                        switchTabs = false;
                        label2.Text = "Select file from list on the side to preview contents. Check and export to save it locally";

                        ResetPreviewFields();

                        treeView1.SelectedNode = null;

                        button2.Enabled = true;

                        statusBar1.Text = "Ready";
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            loadArchive("");
        }

        private string NormalizeTreePath(string path)
        {
            return Regex.Replace(Regex.Replace(path, "^/+", ""), " [(].+[)]$", "");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                List<string> exportFilesList = new List<string>();
                foreach (TreeNode node in treeView1.Nodes.All())
                {
                    if (node.Checked && fileList.ContainsKey(NormalizeTreePath(node.FullPath)))
                    {
                        exportFilesList.Add(NormalizeTreePath(node.FullPath));
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
                    button3.Enabled = true;
                    progressBar1.Value = 0;
                    label1.PerformSafely(() => label1.Text = "");
                    
                    opration = new Thread(() => exportFiles(exportFilesList, folderBrowserDialog));
                    opration.Start();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            operationEnabled = false;
        }

        private void RpaExplorer_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void RpaExplorer_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                loadArchive(((string[]) e.Data.GetData(DataFormats.FileDrop))[0]);
            }
        }

        private void ResetPreviewFields()
        {
            pictureBox1.Image = null;
            textBox2.Text = String.Empty;
            if (videoView1.MediaPlayer != null)
            {
                videoView1.MediaPlayer.Pause();
                videoView1.MediaPlayer.Dispose();
                videoView1.MediaPlayer = null;
            }
            if (mediaVlc != null)
            {
                mediaVlc.Dispose();
                mediaVlc = null;
            }
            if (streamMediaInputVlc != null)
            {
                streamMediaInputVlc.Dispose();
                streamMediaInputVlc = null;
            }
            if (memoryStreamVlc != null)
            {
                memoryStreamVlc.Dispose();
                memoryStreamVlc = null;
            }
        }

        private void treeView1_AfterSelect(object sender, EventArgs e)
        {
            TreeNode selectedNode = new TreeNode();
            bool unsupportedFile = true;
            pictureBox1.Image = null;
            foreach (TreeNode node in treeView1.Nodes.All())
            {
                // Reset fields
                ResetPreviewFields();

                if (node.IsSelected)
                {
                    selectedNode = node;
                }
                
                if (node.IsSelected && fileList.ContainsKey(NormalizeTreePath(node.FullPath)))
                {
                    KeyValuePair<string, object> data = rpaParser.GetPreview(NormalizeTreePath(node.FullPath));
                    if (data.Key == RpaParser.PreviewTypes.Image)
                    {
                        pictureBox1.Image = (Image) data.Value;
                        switchTabs = true;
                        tabControl1.SelectedTab = tabPage1;
                        switchTabs = false;
                        unsupportedFile = false;
                    }
                    else if (data.Key == RpaParser.PreviewTypes.Text)
                    {
                        textBox2.Text = (string) data.Value;
                        switchTabs = true;
                        tabControl1.SelectedTab = tabPage2;
                        switchTabs = false;
                        unsupportedFile = false;
                    }
                    else if (data.Key == RpaParser.PreviewTypes.Audio || data.Key == RpaParser.PreviewTypes.Video)
                    {
                        memoryStreamVlc = new MemoryStream((byte[]) data.Value);
                        streamMediaInputVlc = new StreamMediaInput(memoryStreamVlc);
                        mediaVlc = new Media(libVlc, streamMediaInputVlc);
                        videoView1.MediaPlayer = new MediaPlayer(mediaVlc);
                        videoView1.MediaPlayer.Play();
                        switchTabs = true;
                        tabControl1.SelectedTab = tabPage3;
                        switchTabs = false;
                        unsupportedFile = false;
                    }

                    break;
                }
            }

            if (unsupportedFile)
            {
                switchTabs = true;
                tabControl1.SelectedTab = tabPage0;
                switchTabs = false;
                label2.Text = "Preview is not supported for this file.";
            }
            
            treeView1.SelectedNode = selectedNode;
            treeView1.SelectedNode.EnsureVisible();
            treeView1.Focus();
        }

        private void CheckAllChildNodes(TreeNode node, bool isChecked)
        {
            if (node.Nodes.Count == 0)
            {
                node.Checked = isChecked;
            }
            else
            {
                foreach (TreeNode childNode in node.Nodes)
                {
                    CheckAllChildNodes(childNode, isChecked);
                }
                node.Checked = isChecked;
            }
        }

        private void ParentCheckControl(TreeNode parentNode)
        {
            foreach (TreeNode node in parentNode.Nodes.All())
            {
                if (!node.Checked)
                {
                    parentNode.Checked = false;
                    break;
                }
                parentNode.Checked = true;
            }
            if (parentNode.Parent != null)
            {
                ParentCheckControl(parentNode.Parent);
            }
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Action != TreeViewAction.Unknown)
            {
                if (e.Node.Nodes.Count == 0)
                {
                    if (e.Node.Parent != null)
                    {
                        ParentCheckControl(e.Node.Parent);
                    }
                }
                else if (e.Node.Nodes.Count > 0)
                {
                    CheckAllChildNodes(e.Node, e.Node.Checked);
                }
            }
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (!switchTabs)
            {
                e.Cancel = true;
            }
            switchTabs = false;
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
    
        
    public static class Extensions
    {
        public static IEnumerable<TreeNode> All( this TreeNodeCollection nodes )
        {
            foreach( TreeNode n in nodes )
            {
                yield return n;
                foreach( TreeNode child in n.Nodes.All( ) )
                    yield return child;
            }
        }
    }
}