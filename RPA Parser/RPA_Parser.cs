using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
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
            public const double RPA_1 = 1;
            public const double RPA_2 = 2;
            public const double RPA_3 = 3;
            public const double RPA_3_2 = 3.2;
        }
        
        private class ArchiveMagic
        {
            public const string RPA_1 = ".rpi";
            public const string RPA_2 = "RPA-2.0 ";
            public const string RPA_3 = "RPA-3.0 ";
            public const string RPA_3_2 = "RPA-3.2 ";
        }

        private class CompilationMagic
        {
            public const string RPC_2 = "RENPY RPC2";
        }

        private string _filePath;
        private string _firstLine;
        private FileInfo _fileInfo;
        private double _version;
        private string[] _metadata;
        private long _offset;
        private long _step;
        private SortedDictionary<string,ArchiveIndex> _indexes = new ();

        public class ArchiveIndex
        {
            public long offset = 0;
            public long length = 0;
            public string prefix = String.Empty;
            public string path = String.Empty;
            public string relativePath = String.Empty;
            public bool inArchive = false;
        }

        public class PreviewTypes
        {
            public const string Unknown = "unknown";
            public const string Image = "image";
            public const string Text = "text";
            public const string Video = "video";
            public const string Audio = "audio";
        }

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
            _step = 0xDEADBEEF;
        }

        public void LoadArchive(string filePath)
        {
            _filePath = filePath;
            _fileInfo = GetFileInfo();
            _firstLine = GetFirstLine();
            _version = GetVersion();
            
            if (_version == Version.RPA_2 || _version == Version.RPA_3 || _version == Version.RPA_3_2)
            {
                _metadata = GetMetadata();
                _offset = GetOffset();
                _step = GetStep();
            }

            _indexes = GetIndexes();
        }

        private FileInfo GetFileInfo()
        {
            if (_filePath == String.Empty)
            {
                throw new Exception("No file provided.");
            }

            if (!File.Exists(_filePath))
            {
                throw new Exception("File does not exist.");
            }

            return new FileInfo(_filePath);
        }

        private string GetFirstLine()
        {
            using (var streamReader = new StreamReader(_filePath, Encoding.UTF8))
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

            if (_filePath.EndsWith(ArchiveMagic.RPA_1))
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
            SortedDictionary<string,ArchiveIndex> indexes = new SortedDictionary<string,ArchiveIndex>();
            object unpickledIndexes = new object[]{};
            
            using (BinaryReader reader = new BinaryReader(File.OpenRead(_filePath), Encoding.UTF8))
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

                string key = kvp.Key as string;
                object[] value = (kvp.Value as ArrayList).ToArray()[0] as object[];
                ArchiveIndex index = new ArchiveIndex();
                index.offset = Convert.ToInt64(value.GetValue(0));
                index.length = Convert.ToInt64(value.GetValue(1));
                index.relativePath = key;
                if ((long) value.Length == 3)
                {
                    index.prefix = (string) value.GetValue(2);
                }

                index.inArchive = true;
                indexes.Add(key, index);
            }

            // Deobfuscate index data
            if (_version >= Version.RPA_3)
            {
                foreach (KeyValuePair<string,ArchiveIndex> kvp in indexes)
                {
                    indexes[kvp.Key].offset ^= _step;
                    indexes[kvp.Key].length ^= _step;
                }
            }

            return indexes;
        }

        public SortedDictionary<string, ArchiveIndex> DeepCopyIndex(SortedDictionary<string, ArchiveIndex> index)
        {
            SortedDictionary<string, ArchiveIndex> indexCopy = new SortedDictionary<string, ArchiveIndex>();
            
            foreach (KeyValuePair<string, ArchiveIndex> kvp in index)
            {
                ArchiveIndex archIndex = new ArchiveIndex();
                archIndex.length = kvp.Value.length;
                archIndex.offset = kvp.Value.offset;
                archIndex.path = kvp.Value.path;
                archIndex.prefix = kvp.Value.prefix;
                archIndex.inArchive = kvp.Value.inArchive;
                archIndex.relativePath = kvp.Value.relativePath;
                
                indexCopy.Add(kvp.Key, archIndex);
            }
            
            return indexCopy;
        }

        public SortedDictionary<string, ArchiveIndex> GetFileList()
        {
            return _indexes;
        }

        public void SetFileList(SortedDictionary<string,ArchiveIndex> index)
        {
            _indexes = index;
        }

        public FileInfo GetArchiveInfo()
        {
            return _fileInfo;
        }

        public double GetArchiveVersion()
        {
            return _version;
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
                using (BinaryReader reader = new BinaryReader(File.Open(_filePath, FileMode.Open), Encoding.UTF8))
                {
                    reader.BaseStream.Seek(_indexes[fileName].offset, SeekOrigin.Begin);
                    byte[] prefixData = Encoding.UTF8.GetBytes(_indexes[fileName].prefix);
                    byte[] fileData =
                        reader.ReadBytes((int) _indexes[fileName].length -
                                         _indexes[fileName].prefix.Length); // Exported file max size ~2.14 GB
                    byte[] finalData = new byte[prefixData.Length + fileData.Length];
                    Buffer.BlockCopy(prefixData, 0, finalData, 0, prefixData.Length);
                    Buffer.BlockCopy(fileData, 0, finalData, prefixData.Length, fileData.Length);

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
                finalPath = _fileInfo.DirectoryName + @"\" + fileName;
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

            return _fileInfo.DirectoryName + @"\" + fileName;
        }

        public string SaveArchive(string path, double version, int padding = 0, long step = 0xDEADBEEF)
        {
            if (_version == Version.RPA_1)
            {
                if (!path.EndsWith(".rpi"))
                {
                    path = path + ".rpi";
                }
            }
            else
            {
                if (!path.EndsWith(".rpa"))
                {
                    path = path + ".rpa";
                }
            }

            if (path == _filePath && _filePath != String.Empty)
            {
                throw new Exception("Cannot overwrite same archive that is loaded.");
            }
            
            BuildArchive(path, version, padding, step);

            return path;
        }

        private void BuildArchive(string path, double version, int padding, long step)
        {
            using (Stream stream = File.Open(path, FileMode.Open))
            {
                int archiveOffset = 34; // Default: Version.RPA_3
                switch (version)
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

                    if (padding > 0)
                    {
                        string paddingStr = String.Empty;
                        int paddingLength = rnd.Next(1, padding);

                        while (paddingLength > 0)
                        {
                            paddingStr += Encoding.ASCII.GetString(new byte[] {(byte) rnd.Next(1, 255)});
                            paddingLength--;
                        }

                        byte[] paddingBytes = Encoding.ASCII.GetBytes(paddingStr);
                        archiveOffset += paddingBytes.Length;
                    }

                    if (version == Version.RPA_3)
                    {
                        index.Value.length = content.Length ^ step;
                        index.Value.offset = archiveOffset ^ step;
                    }
                    else if (version == Version.RPA_2)
                    {
                        index.Value.length = content.Length;
                        index.Value.offset = archiveOffset;
                    }

                    stream.Position = archiveOffset;
                    stream.Write(content, 0, content.Length);

                    archiveOffset += content.Length;

                    
                    List<object[]> indexData = new List<object[]>();
                    switch (version) {
                        case Version.RPA_3_2:
                            indexData.Add(new object[] { index.Value.offset, index.Value.length, index.Value.prefix });
                            break;
                        default:
                            indexData.Add(new object[] { index.Value.offset, index.Value.length });
                            break;
                    }

                    indexes.Add(index.Value.relativePath, indexData);
                }

                byte[] pickledIndexes;
                using (Pickler pickler = new Pickler())
                {
                    pickledIndexes = pickler.dumps(indexes);
                }
                byte[] fileCompressed = ZlibStream.CompressBuffer(pickledIndexes);
                
                stream.Position = archiveOffset;
                stream.Write(fileCompressed, 0, fileCompressed.Length);

                string headerContent = String.Empty;

                switch (version)
                {
                    case Version.RPA_3_2:
                        headerContent = ArchiveMagic.RPA_3_2 + archiveOffset.ToString("x").PadLeft(16, '0') + " " + step.ToString("x").PadLeft(8, '0')+ "\n";
                        break;
                    case Version.RPA_3:
                        headerContent = ArchiveMagic.RPA_3 + archiveOffset.ToString("x").PadLeft(16, '0') + " " + step.ToString("x").PadLeft(8, '0') + "\n";
                        break;
                    case Version.RPA_2:
                        headerContent = ArchiveMagic.RPA_2 + archiveOffset.ToString("x").PadLeft(16, '0') + "\n";
                        break;
                }

                byte[] headerContentByte = Encoding.UTF8.GetBytes(headerContent);

                stream.Position = 0;
                stream.Write(headerContentByte, 0, headerContentByte.Length);
            }
        }
    }
}