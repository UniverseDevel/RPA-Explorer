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
        private RpaParser _rpaParser;
        private RpaParser _rpaParserBak;
        private Thread _operation;
        private bool _operationEnabled = true;
        private bool _archiveChanged = false;
        private bool _cancelAdd = false;
        private bool _archiveLoaded = false;
        private SortedDictionary<string, RpaParser.ArchiveIndex> _fileListBackup = new ();
        private readonly Dictionary<string, long> _indexPathSize = new ();
        private readonly List<string> _expandedList = new ();
        private bool _switchTabs = false;
        private readonly LibVLC _libVlc; // https://code.videolan.org/mfkl/libvlcsharp-samples
        private MemoryStream _memoryStreamVlc;
        private StreamMediaInput _streamMediaInputVlc;
        private Media _mediaVlc;

        public RpaExplorer()
        {
            InitializeComponent();
            Core.Initialize();
            _libVlc = new LibVLC("--input-repeat=9999999"); // --input-repeat=X : Loop X times
            videoView1.MediaPlayer = new MediaPlayer(_libVlc);
            videoView1.MediaPlayer.Volume = 80;
            videoView1.MediaPlayer.TimeChanged += videoView1_MediaPlayer_TimeChanged;
            videoView1.BackgroundImage = null;
            SetMediaTimeLabel(0, 0);
            
            ImageList imageList = new ImageList();
            imageList.Images.Add(Resources.treeView1_Image_folder);
            imageList.Images.Add(Resources.treeView1_Image_file);
            imageList.Images.Add(Resources.treeView1_Image_fileChanged);
            treeView1.ImageList = imageList;

            string[] args = Environment.GetCommandLineArgs();
            if (args.Length >= 2)
            {
                LoadArchive(args[1]);
            }
        }

        private void ExportFiles(List<string> exportFilesList, FolderBrowserDialog folderBrowserDialog)
        {
            int counter = 0;
            int jobSize = exportFilesList.Count;
            foreach (string file in exportFilesList)
            {
                counter++;
                int pctProcessed = (int) Math.Ceiling((double) counter / jobSize * 100);
                label1.PerformSafely(() => label1.Text = $@"{counter} / {jobSize}");
                progressBar1.PerformSafely(() => progressBar1.Value = pctProcessed);
                
                statusBar1.PerformSafely(() => statusBar1.Text = Strings_EN.RpaExplorer_exportFiles_Exporting_file__ + file);
                        
                _rpaParser.Extract(file, folderBrowserDialog.SelectedPath);

                if (!_operationEnabled)
                {
                    break;
                }
            }
                    
            button1.PerformSafely(() => button1.Enabled = true);
            button2.PerformSafely(() => button2.Enabled = true);
            button3.PerformSafely(() => button3.Enabled = false);
            progressBar1.PerformSafely(() => progressBar1.Value = 0);
            label1.PerformSafely(() => label1.Text = "");
            statusBar1.PerformSafely(() => statusBar1.Text = Strings_EN.RpaExplorer_Ready);
            
            _operationEnabled = true;
        }

        private void LoadArchive(string openFile, bool ignoreChanges = false)
        {
            if (!ignoreChanges)
            {
                if (CheckIfChanged(Strings_EN.RpaExplorer_LoadArchive_Archive_modified))
                {
                    return;
                }
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = Strings_EN.RpaExplorer_LoadArchive_Select_RenPy_Archive;
                openFileDialog.Filter = Strings_EN.RpaExplorer_RPA_RPI_files;

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
                        statusBar1.Text = Strings_EN.RpaExplorer_LoadArchive_Loading_file__ + openFileDialog.FileName;
                        
                        _expandedList.Clear();
                        _rpaParser = new RpaParser();
                        _rpaParser.LoadArchive(openFileDialog.FileName);

                        GenerateTreeView();

                        _switchTabs = true;
                        tabControl1.SelectedTab = tabPage0;
                        _switchTabs = false;
                        label2.Text = Strings_EN.RpaExplorer_LoadArchive_Choose_file_from_list_on_the_side_to_preview_contents__Check_and_export_to_save_it_locally_;

                        ResetPreviewFields();

                        treeView1.SelectedNode = null;

                        button2.Enabled = true;
                        button6.Enabled = true;
                        button7.Enabled = true;

                        statusBar1.Text = Strings_EN.RpaExplorer_Ready;
                    }

                    _archiveChanged = false;
                }
            }
        }

        private void GenerateTreeView()
        {
            _indexPathSize.Clear();
            treeView1.Nodes.Clear();
            TreeNode root = new TreeNode();
            TreeNode node;
            root.Name = "";
            root.Text = @"/";
            root.BackColor = Color.SandyBrown;
            root.ImageIndex = 0;
            treeView1.Nodes.Add(root);
            _indexPathSize.Add("", 0);
            string pathBuild;

            // Folders first
            foreach (KeyValuePair<string, RpaParser.ArchiveIndex> kvp in _rpaParser.Index)
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
                        if (pathBuild != kvp.Key)
                        {
                            node = node.Nodes.Add(pathBits, pathBits);
                        }
                    }
                }
            }

            // Files second
            foreach (KeyValuePair<string, RpaParser.ArchiveIndex> kvp in _rpaParser.Index)
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
                        if (!_indexPathSize.ContainsKey(pathBuild))
                        {
                            _indexPathSize.Add(pathBuild, 0);
                        }

                        if (_rpaParser.Index.ContainsKey(kvp.Key))
                        {
                            _indexPathSize[pathBuild] += _rpaParser.Index[kvp.Key].Length;
                        }
                    }
                    else
                    {
                        _indexPathSize[""] += _rpaParser.Index[kvp.Key].Length;
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

                if (_rpaParser.Index.ContainsKey(nodeName))
                {
                    if (!_rpaParser.Index[nodeName].InArchive)
                    {
                        nodeVisuals.ForeColor = Color.Green;
                        nodeVisuals.ImageIndex = 2;

                        MarkChanged(nodeVisuals);
                    }
                }
                
                if (_expandedList.Contains(nodeVisuals.FullPath))
                {
                    nodeVisuals.Expand();
                }
            }

            root.Expand();
            _archiveLoaded = true;

            GenerateArchiveInfo();
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
                    if (node.Checked && _rpaParser.Index.ContainsKey(NormalizeTreePath(node.FullPath)))
                    {
                        exportFilesList.Add(NormalizeTreePath(node.FullPath));
                    }
                }
                
                if (exportFilesList.Count == 0)
                {
                    return;
                }
                
                folderBrowserDialog.SelectedPath = _rpaParser.ArchiveInfo.DirectoryName;
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    button1.Enabled = false;
                    button2.Enabled = false;
                    button3.Enabled = true;
                    progressBar1.Value = 0;
                    label1.PerformSafely(() => label1.Text = "");
                    
                    _operation = new Thread(() => ExportFiles(exportFilesList, folderBrowserDialog));
                    _operation.Start();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            _operationEnabled = false;
        }

        private void ResetPreviewFields()
        {
            pictureBox1.Image = null;
            textBox2.Text = String.Empty;
            if (videoView1.MediaPlayer is {IsPlaying: true})
            {
                videoView1.MediaPlayer.Stop();
            }
            if (_mediaVlc != null)
            {
                _mediaVlc.Dispose();
                _mediaVlc = null;
            }
            if (_streamMediaInputVlc != null)
            {
                _streamMediaInputVlc.Dispose();
                _streamMediaInputVlc = null;
            }
            if (_memoryStreamVlc != null)
            {
                _memoryStreamVlc.Dispose();
                _memoryStreamVlc = null;
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
                
                if (node.IsSelected && _rpaParser.Index.ContainsKey(NormalizeTreePath(node.FullPath)))
                {
                    KeyValuePair<string, object> data = _rpaParser.GetPreview(NormalizeTreePath(node.FullPath));
                    if (data.Key == RpaParser.PreviewTypes.Image)
                    {
                        pictureBox1.Image = (Image) data.Value;
                        _switchTabs = true;
                        tabControl1.SelectedTab = tabPage1;
                        _switchTabs = false;
                        unsupportedFile = false;
                    }
                    else if (data.Key == RpaParser.PreviewTypes.Text)
                    {
                        textBox2.Text = (string) data.Value;
                        _switchTabs = true;
                        tabControl1.SelectedTab = tabPage2;
                        _switchTabs = false;
                        unsupportedFile = false;
                    }
                    else if (data.Key == RpaParser.PreviewTypes.Audio || data.Key == RpaParser.PreviewTypes.Video)
                    {
                        _memoryStreamVlc = new MemoryStream((byte[]) data.Value);
                        _streamMediaInputVlc = new StreamMediaInput(_memoryStreamVlc);
                        _mediaVlc = new Media(_libVlc, _streamMediaInputVlc);
                        SetMediaTimeLabel(_mediaVlc.Duration, 0);
                        if (videoView1.MediaPlayer != null) videoView1.MediaPlayer.Play(_mediaVlc);
                        videoView1.BackgroundImage = data.Key == RpaParser.PreviewTypes.Audio ? Resources.videoView1_BackgroundImage : null;
                        _switchTabs = true;
                        tabControl1.SelectedTab = tabPage3;
                        _switchTabs = false;
                        unsupportedFile = false;
                    }

                    break;
                }
            }

            if (unsupportedFile)
            {
                _switchTabs = true;
                tabControl1.SelectedTab = tabPage0;
                _switchTabs = false;
                label2.Text = Strings_EN.RpaExplorer_PreviewSelectedItem_Preview_is_not_supported_for_this_file_;
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
                    if (e.Node.Parent != null)
                    {
                        ParentCheckControl(e.Node.Parent);
                    }
                }
            }
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (!_switchTabs)
            {
                e.Cancel = true;
            }
            _switchTabs = false;
        }

        private void PlayPauseMedia()
        {
            if (videoView1.MediaPlayer is {IsPlaying: true})
            {
                videoView1.MediaPlayer?.Pause();
                button4.Text = Strings_EN.RpaExplorer_Play;
            }
            else
            {
                videoView1.MediaPlayer?.Play();
                button4.Text = Strings_EN.RpaExplorer_Pause;
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
                {
                    // Ignored for now?
                }
            }
            else
            {
                label3.Text = timeText;
            }
        }

        private void videoView1_MediaPlayer_TimeChanged(object sender, MediaPlayerTimeChangedEventArgs e)
        {
            if (videoView1.MediaPlayer?.Media != null)
            {
                SetMediaTimeLabel(videoView1.MediaPlayer.Media.Duration, e.Time);
            }
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            if (videoView1.MediaPlayer?.Media != null)
            {
                videoView1.MediaPlayer.Volume = trackBar1.Value;
            }
        }

        private void RpaExplorer_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (CheckIfChanged(Strings_EN.RpaExplorer_FormClosing_Archive_modified))
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
            if (CheckIfChanged(Strings_EN.RpaExplorer_CreateNewArchive_Archive_modified))
            {
                return;
            }
            
            _expandedList.Clear();
            _rpaParser = new RpaParser();
            
            GenerateTreeView();

            _switchTabs = true;
            tabControl1.SelectedTab = tabPage0;
            _switchTabs = false;
            label2.Text = Strings_EN.RpaExplorer_LoadArchive_Choose_file_from_list_on_the_side_to_preview_contents__Check_and_export_to_save_it_locally_;

            ResetPreviewFields();

            treeView1.SelectedNode = null;

            button2.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;

            statusBar1.Text = Strings_EN.RpaExplorer_Ready;
            
            _archiveChanged = true;
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
            foreach (KeyValuePair<string, RpaParser.ArchiveIndex> kvp in _rpaParser.Index)
            {
                if (!kvp.Value.InArchive)
                {
                    unsavedCount++;
                }

                if (selectedPath == kvp.Key)
                {
                    selectedSize = kvp.Value.Length;
                }
            }

            if (_indexPathSize.ContainsKey(selectedPath))
            {
                selectedSize = _indexPathSize[selectedPath];
            }

            if (!_rpaParser.CheckVersion(_rpaParser.ArchiveVersion, RpaParser.Version.Unknown))
            {
                archiveInfo += "Archive version: " + _rpaParser.ArchiveVersion + Environment.NewLine;
                archiveInfo += "Archive file location: " + _rpaParser.ArchiveInfo.FullName + Environment.NewLine;
                archiveInfo += "Archive file size: " + PrettySize.Format(_indexPathSize[""]) + Environment.NewLine;
            }

            archiveInfo += "Objects count: " + _rpaParser.Index.Count + Environment.NewLine;
            archiveInfo += "Unsaved objects count: " + unsavedCount + Environment.NewLine;

            if (selectedSize != -1)
            {
                archiveInfo += "Object size: " + PrettySize.Format(selectedSize) + Environment.NewLine;
            }

            textBox1.Text = archiveInfo.Trim();
        }

        private bool CheckIfChanged(string message)
        {
            if (message != String.Empty && _archiveChanged)
            {
                if (MessageBox.Show(message, Strings_EN.RpaExplorer_CheckIfChanged_Archive_modified, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                {
                    return true;
                }
                
                return false;
            }
            
            return _archiveChanged;
        }

        private void treeView1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                _fileListBackup.Clear();
                _fileListBackup = _rpaParser.DeepCopyIndex(_rpaParser.Index);
                _cancelAdd = false;
                
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

                if (!_cancelAdd)
                {
                    _rpaParser.Index = _rpaParser.DeepCopyIndex(_fileListBackup);
                }

                _fileListBackup.Clear();

                GenerateTreeView();
            }
        }

        private void AddPathToIndex(string path, string originalPath)
        {
            _archiveChanged = true;
            RpaParser.ArchiveIndex index = new RpaParser.ArchiveIndex();
            index.InArchive = false;
                    
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

            if (File.Exists(path) && !_cancelAdd)
            {
                index.Path = path.Replace(@"\", "/");
                index.RelativePath = index.Path.Replace(originalPath.Replace(@"\", "/") + "/", String.Empty);
                index.Length = new FileInfo(path).Length;
                index.InArchive = false;
                if (_fileListBackup.ContainsKey(index.RelativePath))
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
                    
                    _fileListBackup.Remove(index.RelativePath);
                }
                _fileListBackup.Add(index.RelativePath, index);
            }
        }

        private void treeView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && _archiveLoaded)
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
                if (node.Checked && _rpaParser.Index.ContainsKey(NormalizeTreePath(node.FullPath)))
                {
                    _rpaParser.Index.Remove(NormalizeTreePath(node.FullPath));
                    changed = true;
                }
            }

            if (changed)
            {
                _archiveChanged = true;
                GenerateTreeView();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (_rpaParser.Index.Count == 0)
            {
                MessageBox.Show(Strings_EN.RpaExplorer_button7_Click_Archive_does_not_contain_any_files__cannot_save_empty_archive_,Strings_EN.RpaExplorer_button7_Click_Empty_archive, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            statusBar1.Text = Strings_EN.RpaExplorer_button7_Click_Saving_archive___;

            string initPath = System.Reflection.Assembly.GetEntryAssembly()?.Location;
            if (_rpaParser.ArchiveInfo?.DirectoryName != null)
            {
                initPath = _rpaParser.ArchiveInfo.DirectoryName;
            }
            
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = Strings_EN.RpaExplorer_button7_Click_Save_RenPy_Archive;
            saveFileDialog.Filter = Strings_EN.RpaExplorer_RPA_RPI_files;
            saveFileDialog.InitialDirectory = initPath;
            saveFileDialog.CheckFileExists = false;
            saveFileDialog.CheckPathExists = true;
            
            ArchiveSave options = new ArchiveSave(_rpaParser);
            options.ShowDialog();

            if (!_rpaParser.OptionsConfirmed)
            {
                return;
            }
            
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                _rpaParserBak = _rpaParser;
                try
                {
                    string saveName = _rpaParser.SaveArchive(saveFileDialog.FileName);
                    try
                    {
                        LoadArchive(saveName, true);
                    }
                    catch
                    {
                        _rpaParser = _rpaParserBak;
                        MessageBox.Show(
                            Strings_EN.RpaExplorer_button7_Click_Loading_archive_after_saving_failed_which_indicates_corrupted_archive__Attempting_to_restore_changes_from_before_save_,
                            Strings_EN.RpaExplorer_button7_Click_Save_failed, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    _rpaParser = _rpaParserBak;
                    MessageBox.Show(
                        ex.Message,
                        Strings_EN.RpaExplorer_button7_Click_Save_failed, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                GenerateTreeView();
            }
        }

        private void treeView1_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (_expandedList.Contains(e.Node.FullPath))
            {
                _expandedList.Remove(e.Node.FullPath);
            }
        }

        private void treeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            _expandedList.Add(e.Node.FullPath);
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