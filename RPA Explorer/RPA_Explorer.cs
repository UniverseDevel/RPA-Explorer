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
        private RpaParser rpaParserBak;
        private Thread opration;
        private bool operationEnabled = true;
        private bool archiveChanged = false;
        private bool cancelAdd = false;
        private bool archiveLoaded = false;
        private SortedDictionary<string, RpaParser.ArchiveIndex> fileListBackup = new ();
        private Dictionary<string, long> indexPathSize = new ();
        private List<string> expandedList = new ();
        private string[] args;
        private bool switchTabs = false;
        private LibVLC libVlc; // https://code.videolan.org/mfkl/libvlcsharp-samples
        private MemoryStream memoryStreamVlc;
        private StreamMediaInput streamMediaInputVlc;
        private Media mediaVlc;
        private System.ComponentModel.ComponentResourceManager resources = new (typeof(Global));

        public RpaExplorer()
        {
            InitializeComponent();
            Core.Initialize();
            libVlc = new LibVLC("--input-repeat=9999999"); // --input-repeat=X : Loop X times
            videoView1.MediaPlayer = new MediaPlayer(libVlc);
            videoView1.MediaPlayer.Volume = 80;
            videoView1.MediaPlayer.TimeChanged += videoView1_MediaPlayer_TimeChanged;
            videoView1.BackgroundImage = null;
            SetMediaTimeLabel(0, 0);
            
            ImageList imageList = new ImageList();
            imageList.Images.Add((Image) resources.GetObject("treeView1.Image.folder"), Color.Transparent);
            imageList.Images.Add((Image) resources.GetObject("treeView1.Image.file"), Color.Transparent);
            imageList.Images.Add((Image) resources.GetObject("treeView1.Image.fileChanged"), Color.Transparent);
            treeView1.ImageList = imageList;

            args = Environment.GetCommandLineArgs();
            if (args.Length >= 2)
            {
                LoadArchive(args[1]);
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

        private void LoadArchive(string openFile, bool ignoreChanges = false)
        {
            if (!ignoreChanges)
            {
                if (CheckIfChanged("Archive was modified, do you really want to load a new one and lose changes?"))
                {
                    return;
                }
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select RenPy Archive";
                openFileDialog.Filter = "RPA/RPI files (*.rpa,*.rpi)|*.rpa;*.rpi";

                DialogResult dialogResult = DialogResult.None;
                if (openFile == String.Empty || openFile == null)
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
                        
                        expandedList.Clear();
                        rpaParser = new RpaParser();
                        rpaParser.LoadArchive(openFileDialog.FileName);

                        GenerateTreeView();

                        switchTabs = true;
                        tabControl1.SelectedTab = tabPage0;
                        switchTabs = false;
                        label2.Text = "Choose file from list on the side to preview contents. Check and export to save it locally.";

                        ResetPreviewFields();

                        treeView1.SelectedNode = null;

                        button2.Enabled = true;
                        button6.Enabled = true;
                        button7.Enabled = true;

                        statusBar1.Text = "Ready";
                    }

                    archiveChanged = false;
                }
            }
        }

        private void GenerateTreeView()
        {
            indexPathSize.Clear();
            treeView1.Nodes.Clear();
            TreeNode root = new TreeNode();
            TreeNode node = root;
            root.Name = "";
            root.Text = "/";
            root.BackColor = Color.SandyBrown;
            root.ImageIndex = 0;
            treeView1.Nodes.Add(root);
            indexPathSize.Add("", 0);
            string pathBuild = String.Empty;
            
            foreach (KeyValuePair<string, RpaParser.ArchiveIndex> kvp in rpaParser._indexes)
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
                        node = node.Nodes.Add(pathBits, pathBits);
                    }

                    if (pathBuild != kvp.Key)
                    {
                        if (!indexPathSize.ContainsKey(pathBuild))
                        {
                            indexPathSize.Add(pathBuild, 0);
                        }

                        if (rpaParser._indexes.ContainsKey(kvp.Key))
                        {
                            indexPathSize[pathBuild] += rpaParser._indexes[kvp.Key].length;
                        }
                    }
                    else
                    {
                        indexPathSize[""] += rpaParser._indexes[kvp.Key].length;
                    }
                }
            }
            
            foreach (TreeNode nodeVisuals in treeView1.Nodes.All())
            {
                string nodeName = NormalizeTreePath(nodeVisuals.FullPath);

                if (nodeName != String.Empty)
                {
                    if (nodeVisuals.Nodes.Count > 0)
                    {
                        nodeVisuals.BackColor = Color.SandyBrown;
                        nodeVisuals.ImageIndex = 0;
                    }
                    else
                    {
                        nodeVisuals.BackColor = Color.Transparent;
                        nodeVisuals.ImageIndex = 1;
                    }
                }

                if (rpaParser._indexes.ContainsKey(nodeName))
                {
                    if (!rpaParser._indexes[nodeName].inArchive)
                    {
                        nodeVisuals.ForeColor = Color.Green;
                        nodeVisuals.ImageIndex = 2;

                        MarkChanged(nodeVisuals);
                    }
                }
                
                if (expandedList.Contains(nodeVisuals.FullPath))
                {
                    nodeVisuals.Expand();
                }
            }

            root.Expand();
            archiveLoaded = true;

            GenerateArchiveInfo();
        }

        private long GetNodeSize(TreeNode node)
        {
            long size = 0;

            if (node.Nodes.Count > 0)
            {
                foreach (TreeNode nodeChild in node.Nodes.All())
                {
                    if (nodeChild.Nodes.Count == 0)
                    {
                        string childName = NormalizeTreePath(nodeChild.FullPath);
                        if (rpaParser._indexes.ContainsKey(childName))
                        {
                            size += rpaParser._indexes[childName].length;
                        }
                    }
                }
            }
            else
            {
                if (node.Nodes.Count == 0)
                {
                    string name = NormalizeTreePath(node.FullPath);
                    if (rpaParser._indexes.ContainsKey(name))
                    {
                        size = rpaParser._indexes[name].length;
                    }
                }
            }

            return size;
        }

        private void MarkChanged(TreeNode node)
        {
            if (node.Parent != null)
            {
                node.Parent.ForeColor = Color.Green;
                MarkChanged(node.Parent);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LoadArchive(String.Empty);
        }

        private string NormalizeTreePath(string path)
        {
            return Regex.Replace(path, "^/+", "");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                List<string> exportFilesList = new List<string>();
                foreach (TreeNode node in treeView1.Nodes.All())
                {
                    if (node.Checked && rpaParser._indexes.ContainsKey(NormalizeTreePath(node.FullPath)))
                    {
                        exportFilesList.Add(NormalizeTreePath(node.FullPath));
                    }
                }
                
                if (exportFilesList.Count == 0)
                {
                    return;
                }
                
                folderBrowserDialog.SelectedPath = rpaParser._archiveInfo.DirectoryName;
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

        private void ResetPreviewFields()
        {
            pictureBox1.Image = null;
            textBox2.Text = String.Empty;
            if (videoView1.MediaPlayer.IsPlaying)
            {
                videoView1.MediaPlayer.Pause();
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
            SetMediaTimeLabel(0, 0);
        }

        private void PreviewSelectedItem()
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
                
                if (node.IsSelected && rpaParser._indexes.ContainsKey(NormalizeTreePath(node.FullPath)))
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
                        SetMediaTimeLabel(mediaVlc.Duration, 0);
                        videoView1.MediaPlayer.Play(mediaVlc);
                        if (data.Key == RpaParser.PreviewTypes.Audio)
                        {
                            videoView1.BackgroundImage = (Image) resources.GetObject("videoView1.BackgroundImage");
                        }
                        else
                        {
                            videoView1.BackgroundImage = null;
                        }
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
            treeView1.SelectedNode.SelectedImageIndex = selectedNode.ImageIndex;
            treeView1.SelectedNode.EnsureVisible();
            treeView1.Focus();
        }

        private void treeView1_AfterSelect(object sender, EventArgs e)
        {
            PreviewSelectedItem();
            GenerateArchiveInfo();
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

        private void PlayPauseMedia()
        {
            if (videoView1.MediaPlayer.IsPlaying)
            {
                videoView1.MediaPlayer.Pause();
                button4.Text = "Play";
            }
            else
            {
                videoView1.MediaPlayer.Play();
                button4.Text = "Pause";
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PlayPauseMedia();
        }

        private void SetMediaTimeLabel(long totalTimeMillies, long currentTimeMillies)
        {
            TimeSpan totalTime = new TimeSpan(0, 0, 0, 0, Convert.ToInt32(totalTimeMillies));
            TimeSpan currentTime = new TimeSpan(0, 0, 0, 0, Convert.ToInt32(currentTimeMillies));
            TimeSpan remainingTime = new TimeSpan(0, 0, 0, 0, Convert.ToInt32(totalTimeMillies - currentTimeMillies));

            string timeFormat = "hh':'mm':'ss'.'f";
            string timeText = currentTime.ToString(timeFormat) + " / " + totalTime.ToString(timeFormat) + " (-" + remainingTime.ToString(timeFormat) + ")";
            if (label3.InvokeRequired)
            {
                try
                {
                    label3.PerformSafely(() => label3.Text = timeText);
                }
                catch
                { }
            }
            else
            {
                label3.Text = timeText;
            }
        }

        private void videoView1_MediaPlayer_TimeChanged(object sender, MediaPlayerTimeChangedEventArgs e)
        {
            SetMediaTimeLabel(videoView1.MediaPlayer.Media.Duration, e.Time);
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            videoView1.MediaPlayer.Volume = trackBar1.Value;
        }

        private void RpaExplorer_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (CheckIfChanged("Archive was modified, do you really want to exit without saving and lose changes?"))
            {
                e.Cancel = true;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            CreateNewArchive();
        }

        private void CreateNewArchive()
        {
            if (CheckIfChanged("Archive was modified, do you really want to create a new one and lose changes?"))
            {
                return;
            }
            
            expandedList.Clear();
            rpaParser = new RpaParser();
            rpaParser.CreateArchive();
            
            GenerateTreeView();

            switchTabs = true;
            tabControl1.SelectedTab = tabPage0;
            switchTabs = false;
            label2.Text = "Choose file from list on the side to preview contents. Check and export to save it locally.";

            ResetPreviewFields();

            treeView1.SelectedNode = null;

            button2.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;

            statusBar1.Text = "Ready";
            
            archiveChanged = true;
        }

        private void GenerateArchiveInfo()
        {
            string archiveInfo = String.Empty;

            string selectedPath = String.Empty;
            foreach (TreeNode node in treeView1.Nodes.All())
            {
                if (node.IsSelected)
                {
                    selectedPath = NormalizeTreePath(node.FullPath);
                }
            }

            long selectedSize = -1;
            int unsavedCount = 0;
            foreach (KeyValuePair<string, RpaParser.ArchiveIndex> kvp in rpaParser._indexes)
            {
                if (!kvp.Value.inArchive)
                {
                    unsavedCount++;
                }

                if (selectedPath == kvp.Key)
                {
                    selectedSize = kvp.Value.length;
                }
            }

            if (indexPathSize.ContainsKey(selectedPath))
            {
                selectedSize = indexPathSize[selectedPath];
            }

            if (rpaParser._version != RpaParser.Version.Unknown)
            {
                archiveInfo += "Archive version: " + rpaParser._version + Environment.NewLine;
                archiveInfo += "Archive file location: " + rpaParser._archiveInfo.FullName + Environment.NewLine;
                archiveInfo += "Archive file size: " + PrettySize.Format(indexPathSize[""]) + Environment.NewLine;
            }

            archiveInfo += "Objects count: " + rpaParser._indexes.Count + Environment.NewLine;
            archiveInfo += "Unsaved objects count: " + unsavedCount + Environment.NewLine;

            if (selectedSize != -1)
            {
                archiveInfo += "Object size: " + PrettySize.Format(selectedSize) + Environment.NewLine;
            }

            textBox1.Text = archiveInfo.Trim();
        }

        private bool CheckIfChanged(string message)
        {
            if (message != String.Empty && archiveChanged)
            {
                if (MessageBox.Show(message, "Archive modified", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                {
                    return true;
                }
                
                return false;
            }
            
            return archiveChanged;
        }

        private void treeView1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                fileListBackup.Clear();
                fileListBackup = rpaParser.DeepCopyIndex(rpaParser._indexes);
                cancelAdd = false;
                
                foreach (string path in (string[]) e.Data.GetData(DataFormats.FileDrop))
                {
                    string originalPath = path;
                    if (Directory.Exists(path))
                    {
                        originalPath = new DirectoryInfo(path).Parent?.FullName;
                    }
                    if (File.Exists(path))
                    {
                        originalPath = new FileInfo(path).DirectoryName;
                    }
                    AddPathToIndex(path, originalPath);
                }

                if (!cancelAdd)
                {
                    rpaParser._indexes = rpaParser.DeepCopyIndex(fileListBackup);
                }

                fileListBackup.Clear();

                GenerateTreeView();
            }
        }

        private void AddPathToIndex(string path, string originalPath)
        {
            archiveChanged = true;
            string key = String.Empty;
            RpaParser.ArchiveIndex index = new RpaParser.ArchiveIndex();
            index.inArchive = false;
                    
            if (Directory.Exists(path))
            {
                foreach (string pathFile in Directory.GetFiles(path))
                {
                    AddPathToIndex(pathFile, originalPath);
                }
                foreach (string pathDir in Directory.GetDirectories(path))
                {
                    AddPathToIndex(pathDir, originalPath);
                }
            }

            if (File.Exists(path) && !cancelAdd)
            {
                index.path = path.Replace(@"\", "/");
                index.relativePath = index.path.Replace(originalPath.Replace(@"\", "/") + "/", String.Empty);
                index.length = new FileInfo(path).Length;
                index.inArchive = false;
                if (fileListBackup.ContainsKey(index.relativePath))
                {
                    /*if (fileListBackup[index.relativePath].inArchive)
                    {
                        DialogResult dialogResult = MessageBox.Show("File '" + index.relativePath + "' exists in archive, do you want to replace it?", "File exists in archive", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                        if (dialogResult == DialogResult.Cancel)
                        {
                            cancelAdd = true;
                            return;
                        }
                        if (dialogResult == DialogResult.No)
                        {
                            return;
                        }
                    }*/
                    
                    fileListBackup.Remove(index.relativePath);
                }
                fileListBackup.Add(index.relativePath, index);
            }
        }

        private void treeView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && archiveLoaded)
            {
                e.Effect = DragDropEffects.Copy;
                return;
            }
            e.Effect = DragDropEffects.None;
        }

        private void tabControl1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
                return;
            }
            e.Effect = DragDropEffects.None;
        }

        private void tabControl1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                LoadArchive(((string[]) e.Data.GetData(DataFormats.FileDrop))[0]);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            bool changed = false;
            foreach (TreeNode node in treeView1.Nodes.All())
            {
                if (node.Checked && rpaParser._indexes.ContainsKey(NormalizeTreePath(node.FullPath)))
                {
                    rpaParser._indexes.Remove(NormalizeTreePath(node.FullPath));
                    changed = true;
                }
            }

            if (changed)
            {
                archiveChanged = true;
                GenerateTreeView();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (rpaParser._indexes.Count == 0)
            {
                MessageBox.Show("Archive does not contain any files, cannot save empty archive.","Empty archive", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            statusBar1.Text = "Saving archive...";

            string initPath = System.Reflection.Assembly.GetEntryAssembly().Location;
            if (rpaParser._archiveInfo?.DirectoryName != null)
            {
                initPath = rpaParser._archiveInfo.DirectoryName;
            }
            
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Save RenPy Archive";
            saveFileDialog.Filter = "RPA/RPI files (*.rpa,*.rpi)|*.rpa;*.rpi";
            saveFileDialog.InitialDirectory = initPath;
            saveFileDialog.CheckFileExists = false;
            saveFileDialog.CheckPathExists = true;
            
            ArchiveSave options = new ArchiveSave(rpaParser);
            options.ShowDialog();

            if (!rpaParser.optionsConfirmed)
            {
                return;
            }
            
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                rpaParserBak = rpaParser;
                try
                {
                    string saveName = rpaParser.SaveArchive(saveFileDialog.FileName);
                    try
                    {
                        LoadArchive(saveName, true);
                    }
                    catch
                    {
                        rpaParser = rpaParserBak;
                        MessageBox.Show(
                            "Loading archive after saving failed which indicates corrupted archive. Attempting to restore changes from before save.",
                            "Save failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    rpaParser = rpaParserBak;
                    MessageBox.Show(
                        ex.Message,
                        "Save failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                GenerateTreeView();
            }
        }

        private void treeView1_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (expandedList.Contains(e.Node.FullPath))
            {
                expandedList.Remove(e.Node.FullPath);
            }
        }

        private void treeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            expandedList.Add(e.Node.FullPath);
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