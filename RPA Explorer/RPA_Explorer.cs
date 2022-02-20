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
        private bool _archiveChanged;
        private bool _cancelAdd;
        private bool _archiveLoaded;
        private SortedDictionary<string, RpaParser.ArchiveIndex> _fileListBackup = new ();
        private readonly Dictionary<string, long> _indexPathSize = new ();
        private readonly List<string> _expandedList = new ();
        private bool _switchTabs;
        private readonly LibVLC _libVlc; // https://code.videolan.org/mfkl/libvlcsharp-samples
        private MemoryStream _memoryStreamVlc;
        private StreamMediaInput _streamMediaInputVlc;
        private Media _mediaVlc;

        private readonly FileInfo _appInfo = new (System.Reflection.Assembly.GetEntryAssembly()?.Location ?? throw new InvalidOperationException());

        private static readonly System.ComponentModel.ComponentResourceManager Lang = new (typeof(Lang));
        private static string _language = "EN";

        public RpaExplorer()
        {
            // Form initiation
            InitializeComponent();
            
            // Language initiation
            toolStripComboBox1.Items.Add("English");
            //toolStripComboBox1.Items.Add("TEST"); // For testing purposes

            string storedLanguage = "English";
            if (File.Exists(_appInfo.DirectoryName + @"/lang.ini"))
            {
                storedLanguage = File.ReadAllText(_appInfo.DirectoryName + @"/lang.ini");
            }

            toolStripComboBox1.SelectedItem = toolStripComboBox1.Items.Contains(storedLanguage) ? storedLanguage : "English";

            LoadLanguage(toolStripComboBox1.SelectedItem.ToString());
            LoadTexts();

            // LibVLC initiation
            Core.Initialize();
            _libVlc = new LibVLC("--input-repeat=9999999"); // --input-repeat=X : Loop X times
            videoView1.MediaPlayer = new MediaPlayer(_libVlc);
            videoView1.MediaPlayer.Volume = 80;
            videoView1.MediaPlayer.TimeChanged += videoView1_MediaPlayer_TimeChanged;
            videoView1.BackgroundImage = null;
            SetMediaTimeLabel(0, 0);
            
            // TreeView styles initiation
            ImageList imageList = new ImageList();
            imageList.Images.Add(Resources.treeView1_Image_folder);
            imageList.Images.Add(Resources.treeView1_Image_file);
            imageList.Images.Add(Resources.treeView1_Image_fileChanged);
            treeView1.ImageList = imageList;

            // Arguments processing
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length >= 2)
            {
                LoadArchive(args[1]);
            }
        }

        private void LoadLanguage(string language)
        {
            switch (language)
            {
                case "English":
                    _language = "EN";
                    break;
                case "TEST":
                    _language = "TEST";
                    break;
                default:
                    language = "English";
                    _language = "EN";
                    break;
            }
            
            File.WriteAllText(_appInfo.DirectoryName + @"/lang.ini", language);
        }

        private void LoadTexts()
        {
            Text = GetText("Explorer_title");
            toolStripTextBox1.Text = GetText("Language");
            button1.Text = GetText("Load_file");
            button2.Text = GetText("Export_checked");
            statusBar1.Text = GetText("Ready");
            button3.Text = GetText("Cancel_operation");
            tabPage0.Text = GetText("None");
            label2.Text = GetText("Usage_instructions_new");
            tabPage1.Text = GetText("Image");
            tabPage2.Text = GetText("Text");
            tabPage3.Text = GetText("Media");
            button4.Text = GetText("Pause");
            button5.Text = GetText("Create_new_archive");
            label4.Text = GetText("File_list");
            button6.Text = GetText("Remove_checked");
            button7.Text = GetText("Save_archive");

            GenerateArchiveInfo();
        }

        internal static string GetText(string name)
        {
            if (Lang.GetObject(_language + "_" + name) != null)
            {
                return Lang.GetObject(_language + "_" + name)?.ToString();
            }

            return " {!!! MISSING TRANSLATION !!!} ";
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
                
                statusBar1.PerformSafely(() => statusBar1.Text = GetText("Exporting_file") + file);
                        
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
            statusBar1.PerformSafely(() => statusBar1.Text = GetText("Ready"));
            
            _operationEnabled = true;
        }

        private void LoadArchive(string openFile, bool ignoreChanges = false)
        {
            if (!ignoreChanges)
            {
                if (CheckIfChanged(GetText("Archive_modified_load")))
                {
                    return;
                }
            }

            using OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = GetText("Load_RenPy_Archive");
            openFileDialog.Filter = GetText("RPA_RPI_files") + @" (*.rpa,*.rpi)|*.rpa;*.rpi";

            DialogResult dialogResult = DialogResult.None;
            if (string.IsNullOrEmpty(openFile))
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
                    statusBar1.Text = GetText("Loading_file") + openFileDialog.FileName;

                    try
                    {
                        _expandedList.Clear();
                        _rpaParser = new RpaParser();
                        _rpaParser.LoadArchive(openFileDialog.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            string.Format(GetText("Load_failed_reason"), ex.Message),
                            GetText("Load_failed"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    GenerateTreeView();

                    _switchTabs = true;
                    tabControl1.SelectedTab = tabPage0;
                    _switchTabs = false;
                    label2.Text = GetText("Usage_instructions_loaded");

                    ResetPreviewFields();

                    treeView1.SelectedNode = null;

                    button2.Enabled = true;
                    button6.Enabled = true;
                    button7.Enabled = true;

                    statusBar1.Text = GetText("Ready");
                }

                _archiveChanged = false;
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

                    node = node.Nodes.ContainsKey(pathBits) ? node.Nodes[pathBits] : node.Nodes.Add(pathBits, pathBits);

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

        private string NormalizeTreePath(string path)
        {
            return Regex.Replace(path, "^/+", "");
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
                label2.Text = GetText("Preview_is_not_supported");
            }
            
            treeView1.SelectedNode = selectedNode;
            treeView1.SelectedNode.SelectedImageIndex = selectedNode.ImageIndex;
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

        private void PlayPauseMedia()
        {
            if (videoView1.MediaPlayer is {IsPlaying: true})
            {
                videoView1.MediaPlayer?.Pause();
                button4.Text = GetText("Play");
            }
            else
            {
                videoView1.MediaPlayer?.Play();
                button4.Text = GetText("Pause");
            }
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

        private void CreateNewArchive()
        {
            if (CheckIfChanged(GetText("Archive_modified_new")))
            {
                return;
            }
            
            _expandedList.Clear();
            _rpaParser = new RpaParser();
            
            GenerateTreeView();

            _switchTabs = true;
            tabControl1.SelectedTab = tabPage0;
            _switchTabs = false;
            label2.Text = GetText("Usage_instructions_loaded");

            ResetPreviewFields();

            treeView1.SelectedNode = null;

            button2.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;

            statusBar1.Text = GetText("Ready");
            
            _archiveChanged = true;
        }

        private void GenerateArchiveInfo()
        {
            string archiveInfo = String.Empty;

            if (_archiveLoaded)
            {
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
                    archiveInfo += GetText("Archive_version") + _rpaParser.ArchiveVersion + Environment.NewLine;
                    archiveInfo += GetText("Archive_file_location") + _rpaParser.ArchiveInfo.FullName +
                                   Environment.NewLine;
                    archiveInfo += GetText("Archive_file_size") + PrettySize.Format(_rpaParser.ArchiveInfo.Length) +
                                   Environment.NewLine;
                    if (_rpaParser.IndexInfo != null)
                    {
                        archiveInfo += GetText("Index_file_location") + _rpaParser.IndexInfo.FullName +
                                       Environment.NewLine;
                        archiveInfo += GetText("Index_file_size") + PrettySize.Format(_rpaParser.IndexInfo.Length) +
                                       Environment.NewLine;
                    }
                }

                archiveInfo += GetText("Objects_count") + _rpaParser.Index.Count + Environment.NewLine;
                archiveInfo += GetText("Unsaved_objects_count") + unsavedCount + Environment.NewLine;

                if (selectedSize != -1)
                {
                    archiveInfo += GetText("Selected_object_size") + PrettySize.Format(selectedSize) +
                                   Environment.NewLine;
                }
            }

            textBox1.Text = archiveInfo.Trim();
        }

        private bool CheckIfChanged(string message)
        {
            if (message != String.Empty && _archiveChanged)
            {
                if (MessageBox.Show(message, GetText("Archive_modified"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                {
                    return true;
                }
                
                return false;
            }
            
            return _archiveChanged;
        }

        private void AddPathToIndex(string path, string originalPath)
        {
            _archiveChanged = true;
            RpaParser.ArchiveIndex index = new RpaParser.ArchiveIndex
            {
                InArchive = false
            };

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
                    /*if (_fileListBackup[index.RelativePath].InArchive)
                    {
                        DialogResult dialogResult = MessageBox.Show(string.Format(GetText("Replace_file"), index.RelativePath), GetText("File_exists"), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                        if (dialogResult == DialogResult.Cancel)
                        {
                            _cancelAdd = true;
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

        private void button1_Click(object sender, EventArgs e)
        {
            LoadArchive(String.Empty);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
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

        private void button3_Click(object sender, EventArgs e)
        {
            _operationEnabled = false;
        }

        private void treeView1_AfterSelect(object sender, EventArgs e)
        {
            PreviewSelectedItem();
            GenerateArchiveInfo();
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

        private void button4_Click(object sender, EventArgs e)
        {
            PlayPauseMedia();
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
            if (CheckIfChanged(GetText("Archive_modified_close")))
            {
                e.Cancel = true;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            CreateNewArchive();
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
                MessageBox.Show(GetText("Empty_archive_save"),GetText("Empty_archive"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            statusBar1.Text = GetText("Saving_archive");

            string initPath = _appInfo?.DirectoryName;
            if (_rpaParser.ArchiveInfo?.DirectoryName != null)
            {
                initPath = _rpaParser.ArchiveInfo.DirectoryName;
            }
            
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = GetText("Save_RenPy_Archive");
            saveFileDialog.Filter = GetText("RPA_RPI_files") + @" (*.rpa,*.rpi)|*.rpa;*.rpi";
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
                    LoadArchive(saveName, true);
                }
                catch (Exception ex)
                {
                    _rpaParser = _rpaParserBak;
                    MessageBox.Show(
                        string.Format(GetText("Save_failed_reason"), ex.Message),
                        GetText("Save_failed"), MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadLanguage(toolStripComboBox1.SelectedItem.ToString());
            LoadTexts();
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