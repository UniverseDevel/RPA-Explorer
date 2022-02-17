using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Ionic.Zlib;
using Razorvine.Pickle;
using WebPWrapper;

namespace RPA_Parser
{
    // Inspired by: https://github.com/Shizmob/rpatool/blob/master/rpatool
    // Inspired by: https://github.com/CensoredUsername/unrpyc
    
    public class RpaParser
    {
        public class Version
        {
            public const double Unknown = -1;
            public const double RPA_1 = 1;
            public const double RPA_2 = 2;
            public const double RPA_3 = 3;
            public const double RPA_3_2 = 3.2;
        }
        
        private class ArchiveMagic
        {
            public const string RPA_1_RPA = ".rpa";
            public const string RPA_1_RPI = ".rpi";
            public const string RPA_2 = "RPA-2.0 ";
            public const string RPA_3 = "RPA-3.0 ";
            public const string RPA_3_2 = "RPA-3.2 ";
        }

        private class CompilationMagic
        {
            public const string RPC_2 = "RENPY RPC2";
        }

        public FileInfo _archiveInfo;
        public FileInfo _indexInfo;
        public double _version;
        public long _offset = 0;
        public int _padding = 0;
        public long _step = 0;
        public bool optionsConfirmed = false;
        public SortedDictionary<string,ArchiveIndex> _indexes = new ();
        
        private string _archivePath;
        private string _indexPath;
        private string _firstLine;
        private string[] _metadata;

        public class Index
        {
            public long offset = 0;
            public long length = 0;
            public string prefix = String.Empty;
        }
        public class ArchiveIndex
        {
            public SortedDictionary<int, Index> indexes = new ();
            public string path = String.Empty;
            public string relativePath = String.Empty;
            public bool inArchive = false;
            public long length = 0;
        }

        public class PreviewTypes
        {
            public const string Unknown = "unknown";
            public const string Image = "image";
            public const string Text = "text";
            public const string Video = "video";
            public const string Audio = "audio";
        }
        
        /*
        RenPy Supports:
        Images: JPEG/JPG, PNG, WEBP, BMP, GIF
        Sound/Music: OPUS, OGG Vorbis, FLAC, WAV, MP3, MP2
        Movies: WEBM, OGG Theora, VP9, VP8, MPEG 41, MPEG 2, MPEG 1
        */

        public string[] imageExtList = {
            ".jpeg",
            ".jpg",
            ".bmp",
            ".tiff",
            ".png",
            ".webp",
            ".exif",
            ".ico",
            ".gif"
        };

        public string[] audioExtList = {
            ".aac",
            ".ac3",
            ".flac",
            ".mp3",
            ".wma",
            ".wav",
            ".ogg",
            ".cpc"
        };

        public string[] videoExtList = {
            ".3gp",
            ".flv",
            ".mov",
            ".mp4",
            ".ogv",
            ".swf",
            ".mpg",
            ".mpeg",
            ".avi",
            ".mkv",
            ".wmv",
            ".webm"
        };

        public string[] textExtList = {
            ".py",
            ".rpy~",
            ".rpy",
            ".txt",
            ".log",
            ".nfo",
            ".htm",
            ".html",
            ".xml",
            ".json",
            ".yaml",
            ".csv"
        };

        public string[] rpycExtList = {
            ".rpyc",
            ".rpymc"
            // ".rpyb" ?
        };

        public RpaParser()
        {
            // Init
        }

        public void CreateArchive()
        {
            _version = Version.Unknown;
            _step = 0xDEADBEEF;
        }

        public void LoadArchive(string filePath)
        {
            _archivePath = filePath;
            _archiveInfo = GetArchiveInfo();
            _firstLine = GetFirstLine();
            _version = CheckSupportedVersion(GetVersion());
            
            if (_version == Version.RPA_2 || _version == Version.RPA_3 || _version == Version.RPA_3_2)
            {
                _metadata = GetMetadata();
                _offset = GetOffset();
                _step = GetStep();
            }
            else if (_version == Version.RPA_1)
            {
                GetIndexAndArchive();
            }

            _indexes = GetIndexes();
        }

        private void GetIndexAndArchive()
        {
            if (_version == Version.RPA_1)
            {
                if (_archivePath.EndsWith(ArchiveMagic.RPA_1_RPA))
                {
                    _indexPath = Regex.Replace(_archivePath, @"\.rpa$", ".rpi", RegexOptions.IgnoreCase);
                }
                if (_archivePath.EndsWith(ArchiveMagic.RPA_1_RPI))
                {
                    _indexPath = _archivePath;
                    _archivePath = Regex.Replace(_archivePath, @"\.rpi$", ".rpa", RegexOptions.IgnoreCase);
                }
                
                _archiveInfo = GetArchiveInfo();
                _indexInfo = GetIndexInfo();
            }
        }

        public double CheckSupportedVersion(double version)
        {
            switch (version)
            {
                case Version.RPA_3_2:
                case Version.RPA_3:
                case Version.RPA_2:
                case Version.RPA_1:
                    // Version is OK
                    break;
                default:
                    throw new Exception("Specified version is not supported.");
            }
            
            return version;
        }

        private FileInfo GetArchiveInfo()
        {
            if (_archivePath == String.Empty)
            {
                throw new Exception("No archive file provided.");
            }

            if (!File.Exists(_archivePath))
            {
                throw new Exception("Archive file does not exist.");
            }

            return new FileInfo(_archivePath);
        }

        private FileInfo GetIndexInfo()
        {
            if (_indexPath == String.Empty)
            {
                throw new Exception("No index file provided.");
            }

            if (!File.Exists(_indexPath))
            {
                throw new Exception("Index file does not exist.");
            }

            return new FileInfo(_indexPath);
        }

        private string GetFirstLine()
        {
            using (var streamReader = new StreamReader(_archivePath, Encoding.UTF8))
            {
                var firstLine = streamReader.ReadLine();

                return firstLine;
            }
        }

        private double GetVersion()
        {
            if (_firstLine.StartsWith(ArchiveMagic.RPA_3_2))
            {
                return 3.2;
            }

            if (_firstLine.StartsWith(ArchiveMagic.RPA_3))
            {
                return 3;
            }

            if (_firstLine.StartsWith(ArchiveMagic.RPA_2))
            {
                return 2;
            }

            if (_archivePath.EndsWith(ArchiveMagic.RPA_1_RPA) || _archivePath.EndsWith(ArchiveMagic.RPA_1_RPI))
            {
                return 1;
            }

            throw new Exception("File is either not valid RenPy Archive or version is not supported.");
        }

        private string[] GetMetadata()
        {
            return _firstLine.Split(' ');
        }

        private long GetOffset()
        {
            return Convert.ToInt64(_metadata[1], 16);
        }

        private long GetStep()
        {
            long step = 0;
            
            if (_version == Version.RPA_3)
            {
                for(var i = 2; i < _metadata.Length; i++)
                {
                    step ^= Convert.ToInt64(_metadata[i], 16);
                }
            }
            else if (_version == Version.RPA_3_2)
            {
                for(var i = 3; i < _metadata.Length; i++)
                {
                    step ^= Convert.ToInt64(_metadata[i], 16);
                }
            }

            return step;
        }
        
        private SortedDictionary<string,ArchiveIndex> GetIndexes()
        {
            SortedDictionary<string,ArchiveIndex> indexList = new SortedDictionary<string,ArchiveIndex>();
            object unpickledIndexes = new object[]{};

            string filePath = _archivePath;
            if (_version == Version.RPA_1)
            {
                filePath = _indexPath;
            }
            
            using (BinaryReader reader = new BinaryReader(File.OpenRead(filePath), Encoding.UTF8))
            {
                if (_version == Version.RPA_2 || _version == Version.RPA_3 || _version == Version.RPA_3_2)
                {
                    reader.BaseStream.Seek(_offset, SeekOrigin.Begin);
                }

                long blockOffset = _offset;
                long blockSize = 2046;
                long payloadSize = reader.BaseStream.Length;
                byte[] fileCompressed = new byte[0];

                while (blockSize > 0)
                {
                    long remaining = payloadSize - blockOffset;
                    if (blockOffset + blockSize > payloadSize)
                    {
                        blockSize = payloadSize - blockOffset;

                        if (blockSize < 0)
                        {
                            blockSize = 0;
                        }
                    }

                    if (blockSize != 0)
                    {
                        byte[] buffer = new byte[blockSize];
                        buffer = reader.ReadBytes((int) blockSize);
                        fileCompressed = fileCompressed.Concat(buffer).ToArray();

                        blockOffset += blockSize;
                        reader.BaseStream.Seek(blockOffset, SeekOrigin.Begin);
                    }
                }


                byte[] fileUncompressed = ZlibStream.UncompressBuffer(fileCompressed);
                using (Unpickler unpickler = new Unpickler())
                {
                    unpickledIndexes = unpickler.loads(fileUncompressed);
                }
            }
            
            // Standardize output
            foreach (DictionaryEntry kvp in (Hashtable) unpickledIndexes)
            {
                if (kvp.Value == null)
                {
                    continue;
                }

                ArchiveIndex indexEntry = new ArchiveIndex();
                indexEntry.relativePath = (string) kvp.Key;
                indexEntry.inArchive = true;
                int counter = 0;
                foreach (object[] value in (ArrayList) kvp.Value) // Version 1 can contain multiple positions of same file?? Need to nest this without [0]
                { 
                    Index index = new Index();
                    index.offset = Convert.ToInt64(value.GetValue(0));
                    index.length = Convert.ToInt64(value.GetValue(1));
                    if ((long) value.Length == 3)
                    {
                        index.prefix = (string) value.GetValue(2);
                    }

                    indexEntry.indexes.Add(counter, index);
                    counter++;
                }
                indexList.Add(indexEntry.relativePath, indexEntry);
            }

            foreach (KeyValuePair<string, ArchiveIndex> kvp in indexList)
            {
                foreach (KeyValuePair<int, Index> kvpI in kvp.Value.indexes)
                {
                    // Deobfuscate index data
                    if (_version >= Version.RPA_3)
                    {
                        kvpI.Value.offset ^= _step;
                        kvpI.Value.length ^= _step;
                    }

                    kvp.Value.length += kvpI.Value.length;
                }
            }

            return indexList;
        }

        public SortedDictionary<string, ArchiveIndex> DeepCopyIndex(SortedDictionary<string, ArchiveIndex> originalIndex)
        {
            SortedDictionary<string, ArchiveIndex> indexCopy = new SortedDictionary<string, ArchiveIndex>();
            
            foreach (KeyValuePair<string, ArchiveIndex> kvp in originalIndex)
            {
                ArchiveIndex archIndex = new ArchiveIndex();
                archIndex.path = kvp.Value.path;
                archIndex.inArchive = kvp.Value.inArchive;
                archIndex.relativePath = kvp.Value.relativePath;
                foreach (KeyValuePair<int, Index> kvpI in kvp.Value.indexes)
                {
                    Index index = new Index();
                    index.length = kvpI.Value.length;
                    index.offset = kvpI.Value.offset;
                    index.prefix = kvpI.Value.prefix;
                
                    archIndex.indexes.Add(kvpI.Key, index);
                }
                
                indexCopy.Add(kvp.Key, archIndex);
            }
            
            return indexCopy;
        }

        public KeyValuePair<string, byte[]> GetPreviewRaw(string fileName)
        {
            KeyValuePair<string, object> data = GetPreview(fileName, true);
            return new KeyValuePair<string, byte[]>(data.Key, (byte[]) data.Value);
        }

        public KeyValuePair<string, object> GetPreview(string fileName, bool returnRaw = false)
        {
            KeyValuePair<string, object> data = new KeyValuePair<string, object>(PreviewTypes.Unknown, null);

            if (!_indexes.ContainsKey(fileName))
            {
                return data;
            }

            FileInfo fileInfo = new FileInfo(fileName);
            byte[] bytes = ExtractData(fileName);

            if (imageExtList.Contains(fileInfo.Extension.ToLower()))
            {
                byte[] magicBytes = new byte[16] ;
                Buffer.BlockCopy(bytes,0, magicBytes,0,16);

                Image image = null;
                if (fileInfo.Extension.ToLower() == ".webp" || Encoding.UTF8.GetString(magicBytes, 0, magicBytes.Length).Contains("WEBP"))
                {
                    using (WebP ww = new WebP())
                    {
                        image = ww.Decode(bytes);
                    }
                }
                else
                {
                    using (MemoryStream ms = new MemoryStream(bytes))
                    {
                        image = Image.FromStream(ms);
                    }
                }

                data = new KeyValuePair<string, object>(PreviewTypes.Image, image);
            }
            else if (textExtList.Contains(fileInfo.Extension.ToLower()))
            {
                data = new KeyValuePair<string, object>(PreviewTypes.Text, NormlizeNewLines(Encoding.UTF8.GetString(bytes, 0, bytes.Length)));
            }
            else if (audioExtList.Contains(fileInfo.Extension.ToLower()))
            {
                data = new KeyValuePair<string, object>(PreviewTypes.Audio, bytes);
            }
            else if (videoExtList.Contains(fileInfo.Extension.ToLower()))
            {
                data = new KeyValuePair<string, object>(PreviewTypes.Video, bytes);
            }
            else if (rpycExtList.Contains(fileInfo.Extension.ToLower()))
            {
                bytes = DeobfuscateRPC(bytes);
                data = new KeyValuePair<string, object>(PreviewTypes.Text, NormlizeNewLines(Encoding.UTF8.GetString(bytes, 0, bytes.Length)));
            }
            else
            {
                data = new KeyValuePair<string, object>(PreviewTypes.Unknown, bytes);
            }

            if (returnRaw)
            {
                if (rpycExtList.Contains(fileInfo.Extension.ToLower()))
                {
                    data = new KeyValuePair<string, object>(data.Key, bytes);
                }
                else
                {
                    data = new KeyValuePair<string, object>(data.Key, bytes);
                }
            }

            return data;
        }

        private string NormlizeNewLines(string text)
        {
            const string winNewLine = "\r\n";
            const string linNewLine = "\n";
            const string macNewLine = "\r";
            
            int countWin = Regex.Matches(text, winNewLine).Count;
            int countLinux = Regex.Matches(text, linNewLine).Count;
            int countMac = Regex.Matches(text, macNewLine).Count;
            
            string newLineSymbol = Environment.NewLine;
            
            if (countWin >= countLinux && countWin >= countMac)
            {
                newLineSymbol = winNewLine;
            }
            else if (countLinux >= countWin && countLinux >= countMac)
            {
                newLineSymbol = linNewLine;
            }
            else if (countMac >= countWin && countMac >= countLinux)
            {
                newLineSymbol = macNewLine;
            }

            text = text.Replace(newLineSymbol, Environment.NewLine);
            
            return text;
        }

        private byte[] DeobfuscateRPC(byte[] fileData)
        {
            string fileText = String.Empty;
            byte[] magicBytes = new byte[10] ;
            Buffer.BlockCopy(fileData,0, magicBytes,0,10);
            
            if (Encoding.UTF8.GetString(magicBytes, 0, magicBytes.Length).StartsWith(CompilationMagic.RPC_2))
            {
                // TODO: rpyc parser?
                fileText = Encoding.UTF8.GetString(fileData, 0, fileData.Length);
                fileText = "Not yet supported."; // TODO: remove when implemented
            }
            else
            {
                // Legacy files might not have header so we should assume that its a zlib package unless zlib fails to do anything with it
                try
                {
                    // TODO: handle legacy files
                }
                catch
                {
                    throw new Exception("File is either not valid RenPy Compilation or version is not supported.");
                }
            }
            
            return Encoding.UTF8.GetBytes(fileText);
        }

        public byte[] ExtractData(string fileName)
        {
            if (!_indexes.ContainsKey(fileName))
            {
                throw new Exception("Specified file does not exist in RenPy Archive.");
            }

            if (_indexes[fileName].inArchive)
            {
                using (BinaryReader reader = new BinaryReader(File.Open(_archivePath, FileMode.Open), Encoding.UTF8))
                {
                    byte[] finalData = { };

                    foreach (KeyValuePair<int, Index> kvpI in _indexes[fileName].indexes)
                    {
                        reader.BaseStream.Seek(kvpI.Value.offset, SeekOrigin.Begin);
                        byte[] prefixData = Encoding.UTF8.GetBytes(kvpI.Value.prefix);
                        byte[] fileData = reader.ReadBytes((int) kvpI.Value.length - kvpI.Value.prefix.Length); // Exported file max size ~2.14 GB
                        byte[] partData = new byte[finalData.Length + prefixData.Length + fileData.Length];
                        Buffer.BlockCopy(finalData, 0, partData, 0, finalData.Length);
                        Buffer.BlockCopy(prefixData, 0, partData, finalData.Length, prefixData.Length);
                        Buffer.BlockCopy(fileData, 0, partData, finalData.Length + prefixData.Length, fileData.Length);
                        finalData = partData;
                    }

                    return finalData;
                }
            }
            
            return File.ReadAllBytes(_indexes[fileName].path);
        }

        public string Extract(string fileName, string exportPath)
        {
            byte[] finalData = ExtractData(fileName);
            string finalPath = String.Empty;
            if (exportPath.Trim() == null || exportPath.Trim() == String.Empty)
            {
                finalPath = _archiveInfo.DirectoryName + @"\" + fileName;
            }
            else
            {
                if (!Directory.Exists(exportPath.Trim()))
                {
                    throw new Exception("Selected export path does not exist.");
                }
                finalPath = exportPath.Trim() + @"\" + fileName;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(finalPath));
            File.WriteAllBytes(finalPath, finalData);

            return _archiveInfo.DirectoryName + @"\" + fileName;
        }

        public string SaveArchive(string archivePath)
        {
            if (archivePath.EndsWith(".rpi"))
            {
                archivePath = Regex.Replace(archivePath, @"\.rpi$", ".rpa", RegexOptions.IgnoreCase);
            }
            
            if (!archivePath.EndsWith(".rpa"))
            {
                archivePath += ".rpa";
            }

            if (archivePath == _archivePath && _archivePath != String.Empty)
            {
                throw new Exception("Cannot overwrite same archive file that is loaded.");
            }

            string indexPath = Regex.Replace(archivePath, @"\.rpa$", ".rpi", RegexOptions.IgnoreCase);

            if (indexPath == _indexPath && _indexPath != String.Empty)
            {
                throw new Exception("Cannot overwrite same index file that is loaded.");
            }
            
            BuildArchive(archivePath, indexPath);

            return archivePath;
        }

        private void BuildArchive(string archivePath, string indexPath)
        {
            if (!File.Exists(archivePath))
            {
                File.WriteAllBytes(archivePath, new byte[] { });
            }
            
            using (Stream stream = File.Open(archivePath, FileMode.Truncate))
            {
                int archiveOffset = 34; // Default: Version.RPA_3
                switch (_version)
                {
                    case Version.RPA_3_2:
                        archiveOffset = 34;
                        break;
                    case Version.RPA_3:
                        archiveOffset = 34;
                        break;
                    case Version.RPA_2:
                        archiveOffset = 25;
                        break;
                    case Version.RPA_1:
                        archiveOffset = 0;
                        break;
                    default:
                        throw new Exception("Specified version is not supported.");
                }

                stream.Position = archiveOffset;
                
                Random rnd = new Random();
                
                // Update indexes
                Hashtable indexes = new Hashtable();
                foreach (KeyValuePair<string, ArchiveIndex> index in _indexes)
                {
                    byte[] content = ExtractData(index.Key);

                    if (_padding > 0)
                    {
                        string paddingStr = String.Empty;
                        int paddingLength = rnd.Next(1, _padding);

                        while (paddingLength > 0)
                        {
                            paddingStr += Encoding.ASCII.GetString(new [] {(byte) rnd.Next(1, 255)});
                            paddingLength--;
                        }

                        byte[] paddingBytes = Encoding.ASCII.GetBytes(paddingStr);
                        archiveOffset += paddingBytes.Length;
                    }

                    stream.Position = archiveOffset;
                    stream.Write(content, 0, content.Length);

                    List<object[]> indexData = new List<object[]>();
                    if (_version == Version.RPA_3 || _version == Version.RPA_3_2)
                    {
                        indexData.Add(new object[] {archiveOffset ^ _step, content.Length ^ _step, ""}); // Last is prefix
                    }
                    else
                    {
                        indexData.Add(new object[] {archiveOffset, content.Length});
                    }

                    archiveOffset += content.Length;

                    indexes.Add(index.Value.relativePath, indexData);
                }

                byte[] pickledIndexes;
                using (Pickler pickler = new Pickler())
                {
                    pickledIndexes = pickler.dumps(indexes);
                }
                byte[] fileCompressed = ZlibStream.CompressBuffer(pickledIndexes);

                if (_version != Version.RPA_1)
                {
                    stream.Position = archiveOffset;
                    stream.Write(fileCompressed, 0, fileCompressed.Length);

                    string headerContent = String.Empty;

                    switch (_version)
                    {
                        case Version.RPA_3_2:
                            headerContent = ArchiveMagic.RPA_3_2 + archiveOffset.ToString("x").PadLeft(16, '0') + " " +
                                            _step.ToString("x").PadLeft(8, '0') + "\n";
                            break;
                        case Version.RPA_3:
                            headerContent = ArchiveMagic.RPA_3 + archiveOffset.ToString("x").PadLeft(16, '0') + " " +
                                            _step.ToString("x").PadLeft(8, '0') + "\n";
                            break;
                        case Version.RPA_2:
                            headerContent = ArchiveMagic.RPA_2 + archiveOffset.ToString("x").PadLeft(16, '0') + "\n";
                            break;
                    }

                    byte[] headerContentByte = Encoding.UTF8.GetBytes(headerContent);

                    stream.Position = 0;
                    stream.Write(headerContentByte, 0, headerContentByte.Length);
                }
                else
                {
                    File.WriteAllBytes(indexPath, fileCompressed);
                }
            }
        }
    }
}