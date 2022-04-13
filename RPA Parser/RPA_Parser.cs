using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
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
    // Inspired by: https://github.com/Shizmob/rpatool
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
        
        private class RPCMagic
        {
            public const string RPC_2 = "RENPY RPC2";
        }

        public FileInfo ArchiveInfo;
        public FileInfo IndexInfo;
        public double ArchiveVersion = Version.Unknown;
        public int Padding = 0;
        public long Step = 0xDEADBEEF;
        public bool OptionsConfirmed = false;
        public SortedDictionary<string,ArchiveIndex> Index = new ();
        
        private long _offset;
        private string _archivePath;
        private string _indexPath;
        private string _firstLine;
        private string[] _metadata;

        public class Tuples
        {
            public long Offset;
            public long Length;
            public string Prefix = String.Empty;
        }
        
        public class ArchiveIndex
        {
            public readonly SortedDictionary<int, Tuples> Tuples = new ();
            public string FullPath = String.Empty;
            public string TreePath = String.Empty;
            public string ParentPath = String.Empty;
            public bool InArchive;
            public long Length;
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

        public readonly string[] ImageExtList = {
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

        public readonly string[] AudioExtList = {
            ".aac",
            ".ac3",
            ".flac",
            ".mp3",
            ".wma",
            ".wav",
            ".ogg",
            ".cpc"
        };

        public readonly string[] VideoExtList = {
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

        public readonly string[] TextExtList = {
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

        public readonly string[] CodeExtList = {
            ".rpyc~",
            ".rpyc",
            ".rpymc~",
            ".rpymc"
        };

        public void LoadArchive(string filePath)
        {
            _archivePath = filePath;
            GetIndexAndArchive();
            ArchiveInfo = GetArchiveInfo();
            _firstLine = GetFirstLine();
            ArchiveVersion = CheckSupportedVersion(GetVersion());
            
            if (CheckVersion(ArchiveVersion, Version.RPA_2) || CheckVersion(ArchiveVersion, Version.RPA_3) || CheckVersion(ArchiveVersion, Version.RPA_3_2))
            {
                _metadata = GetMetadata();
                _offset = GetOffset();
                Step = GetStep();
            }
            else if (CheckVersion(ArchiveVersion, Version.RPA_1))
            {
                IndexInfo = GetIndexInfo();
            }

            Index = GetIndexes();
        }

        public bool CheckVersion(double version, double check)
        {
            double difference = version - check;
            if (difference == 0)
            {
                return true;
            }

            return false;
        }

        private void GetIndexAndArchive()
        {
            if (_archivePath.ToLower().EndsWith(ArchiveMagic.RPA_1_RPA))
            {
                _indexPath = Regex.Replace(_archivePath, @"\.rpa$", ".rpi", RegexOptions.IgnoreCase);
            }
            if (_archivePath.ToLower().EndsWith(ArchiveMagic.RPA_1_RPI))
            {
                _indexPath = _archivePath;
                _archivePath = Regex.Replace(_archivePath, @"\.rpi$", ".rpa", RegexOptions.IgnoreCase);
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
            using StreamReader streamReader = new StreamReader(_archivePath, Encoding.UTF8);
            return streamReader.ReadLine();
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

            if (_archivePath.ToLower().EndsWith(ArchiveMagic.RPA_1_RPA) || _archivePath.ToLower().EndsWith(ArchiveMagic.RPA_1_RPI))
            {
                GetIndexAndArchive();
                if (File.Exists(_archivePath) && File.Exists(_indexPath))
                {
                    return 1;
                }
            }

            throw new Exception("File is either not valid RenPy Archive or version is not recognized.");
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
            
            if (CheckVersion(ArchiveVersion, Version.RPA_3))
            {
                for(int i = 2; i < _metadata.Length; i++)
                {
                    step ^= Convert.ToInt64(_metadata[i], 16);
                }
            }
            else if (CheckVersion(ArchiveVersion, Version.RPA_3_2))
            {
                for(int i = 3; i < _metadata.Length; i++)
                {
                    step ^= Convert.ToInt64(_metadata[i], 16);
                }
            }

            return step;
        }
        
        private SortedDictionary<string,ArchiveIndex> GetIndexes()
        {
            SortedDictionary<string,ArchiveIndex> indexList = new SortedDictionary<string,ArchiveIndex>();
            object unpickledIndexes;

            string filePath = _archivePath;
            if (CheckVersion(ArchiveVersion, Version.RPA_1))
            {
                filePath = _indexPath;
            }
            
            using (BinaryReader reader = new BinaryReader(File.OpenRead(filePath), Encoding.UTF8))
            {
                if (CheckVersion(ArchiveVersion, Version.RPA_2) || CheckVersion(ArchiveVersion, Version.RPA_3) || CheckVersion(ArchiveVersion, Version.RPA_3_2))
                {
                    reader.BaseStream.Seek(_offset, SeekOrigin.Begin);
                }

                long blockOffset = _offset;
                long blockSize = 2046;
                long payloadSize = reader.BaseStream.Length;
                byte[] fileCompressed = { };

                while (blockSize > 0)
                {
                    //long remaining = payloadSize - blockOffset;
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
                        byte[] buffer = reader.ReadBytes((int) blockSize);
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

                ArchiveIndex indexEntry = new ArchiveIndex
                {
                    TreePath = (string) kvp.Key,
                    ParentPath = Path.GetDirectoryName((string) kvp.Key),
                    InArchive = true
                };
                int counter = 0;
                foreach (object[] value in (ArrayList) kvp.Value)
                { 
                    Tuples index = new Tuples
                    {
                        Offset = Convert.ToInt64(value.GetValue(0)),
                        Length = Convert.ToInt64(value.GetValue(1))
                    };
                    if ((long) value.Length == 3)
                    {
                        index.Prefix = (string) value.GetValue(2);
                    }

                    indexEntry.Tuples.Add(counter, index);
                    counter++;
                }
                indexList.Add(indexEntry.TreePath, indexEntry);
            }

            foreach (KeyValuePair<string, ArchiveIndex> kvp in indexList)
            {
                foreach (KeyValuePair<int, Tuples> kvpI in kvp.Value.Tuples)
                {
                    // Deobfuscate index data
                    if (ArchiveVersion >= Version.RPA_3)
                    {
                        kvpI.Value.Offset ^= Step;
                        kvpI.Value.Length ^= Step;
                    }

                    kvp.Value.Length += kvpI.Value.Length;
                }
            }

            return indexList;
        }

        public SortedDictionary<string, ArchiveIndex> DeepCopyIndex(SortedDictionary<string, ArchiveIndex> originalIndex)
        {
            SortedDictionary<string, ArchiveIndex> indexCopy = new SortedDictionary<string, ArchiveIndex>();
            
            foreach (KeyValuePair<string, ArchiveIndex> kvp in originalIndex)
            {
                ArchiveIndex archIndex = new ArchiveIndex
                {
                    FullPath = kvp.Value.FullPath,
                    InArchive = kvp.Value.InArchive,
                    TreePath = kvp.Value.TreePath,
                    ParentPath = kvp.Value.ParentPath,
                    Length = kvp.Value.Length
                };
                
                foreach (KeyValuePair<int, Tuples> kvpI in kvp.Value.Tuples)
                {
                    Tuples index = new Tuples
                    {
                        Length = kvpI.Value.Length,
                        Offset = kvpI.Value.Offset,
                        Prefix = kvpI.Value.Prefix
                    };
                    
                    archIndex.Tuples.Add(kvpI.Key, index);
                }
                
                indexCopy.Add(kvp.Key, archIndex);
            }
            
            return indexCopy;
        }

        public string ParseRPYC(byte[] code)
        {
            string decompiled = String.Empty;

            if (Encoding.UTF8.GetString(code).StartsWith(RPCMagic.RPC_2))
            {
                long blockOffset = 10;
                Dictionary<int, byte[]> chunkList = new Dictionary<int, byte[]>();

                while (true)
                {
                    byte[] chunkPart = new byte[12];
                    Buffer.BlockCopy(code, (int) blockOffset, chunkPart, 0, 12);
                    object[] structData = StructConverter.Unpack("III", chunkPart); // slot, start, length
                    if ((int) structData[0] == 0)
                    {
                        break;
                    }
                    blockOffset += 12;
                    
                    byte[] chunk = new byte[(int) structData[2]];
                    Buffer.BlockCopy(code, (int) structData[1], chunk, 0, (int) structData[2]);
                    
                    chunkList.Add((int) structData[0], chunk);
                }

                byte[] fileUncompressed;
                try
                {
                    fileUncompressed = ZlibStream.UncompressBuffer(chunkList[1]);
                }
                catch (ZlibException ex)
                {
                    throw new Exception("Parsed slot 1 is not Zlib BLOB. " + ex.Message);
                }

                if (!Encoding.UTF8.GetString(fileUncompressed).EndsWith("."))
                {
                    throw new Exception("Parsed uncompressed slot 1 is not simple pickle.");
                }

                // TODO: pickletools.dis => disassembly, seems like there is no out of the box alternative for this in C#

                //decompiled = Encoding.UTF8.GetString(fileUncompressed);
            }

            throw new NotImplementedException(); // TODO: remove when done
            return decompiled;
        }

        public KeyValuePair<string, byte[]> GetPreviewRaw(string fileName)
        {
            KeyValuePair<string, object> data = GetPreview(fileName, true);
            return new KeyValuePair<string, byte[]>(data.Key, (byte[]) data.Value);
        }

        public KeyValuePair<string, object> GetPreview(string fileName, bool returnRaw = false)
        {
            KeyValuePair<string, object> data = new KeyValuePair<string, object>(PreviewTypes.Unknown, null);

            if (!Index.ContainsKey(fileName))
            {
                return data;
            }

            FileInfo fileInfo = new FileInfo(fileName);
            byte[] bytes = ExtractData(fileName);

            if (ImageExtList.Contains(fileInfo.Extension.ToLower()))
            {
                byte[] magicBytes = new byte[16];
                Buffer.BlockCopy(bytes,0, magicBytes,0,16);

                Image image;
                if (fileInfo.Extension.ToLower() == ".webp" || Encoding.UTF8.GetString(magicBytes, 0, magicBytes.Length).Contains("WEBP"))
                {
                    image = new WebP().Decode(bytes);
                }
                else
                {
                    image = Image.FromStream(new MemoryStream(bytes));
                }

                data = new KeyValuePair<string, object>(PreviewTypes.Image, image);
            }
            else if (TextExtList.Contains(fileInfo.Extension.ToLower()))
            {
                data = new KeyValuePair<string, object>(PreviewTypes.Text, NormalizeNewLines(Encoding.UTF8.GetString(bytes, 0, bytes.Length)));
            }
            else if (CodeExtList.Contains(fileInfo.Extension.ToLower()))
            {
                string decompiledString = String.Empty;
                
                try
                {
                    decompiledString = ParseRPYC(bytes);
                }
                catch
                {
                    // Ignored
                }

                if (decompiledString == String.Empty)
                {
                    data = new KeyValuePair<string, object>(PreviewTypes.Unknown, bytes);
                }
                else
                {
                    data = new KeyValuePair<string, object>(PreviewTypes.Text, decompiledString);
                }
            }
            else if (AudioExtList.Contains(fileInfo.Extension.ToLower()))
            {
                data = new KeyValuePair<string, object>(PreviewTypes.Audio, bytes);
            }
            else if (VideoExtList.Contains(fileInfo.Extension.ToLower()))
            {
                data = new KeyValuePair<string, object>(PreviewTypes.Video, bytes);
            }
            else
            {
                data = new KeyValuePair<string, object>(PreviewTypes.Unknown, bytes);
            }

            if (returnRaw)
            {
                data = new KeyValuePair<string, object>(data.Key, bytes);
            }

            return data;
        }

        private string NormalizeNewLines(string text)
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

        public byte[] ExtractData(string fileName)
        {
            if (!Index.ContainsKey(fileName))
            {
                throw new Exception("Specified file does not exist in RenPy Archive.");
            }

            if (Index[fileName].InArchive)
            {
                using BinaryReader reader = new BinaryReader(File.OpenRead(_archivePath), Encoding.UTF8);
                byte[] finalData = { };

                foreach (KeyValuePair<int, Tuples> kvpI in Index[fileName].Tuples)
                {
                    reader.BaseStream.Seek(kvpI.Value.Offset, SeekOrigin.Begin);
                    byte[] prefixData = Encoding.UTF8.GetBytes(kvpI.Value.Prefix);
                    byte[] fileData = reader.ReadBytes((int) kvpI.Value.Length - kvpI.Value.Prefix.Length); // Exported file max size ~2.14 GB
                    byte[] partData = new byte[finalData.Length + prefixData.Length + fileData.Length];
                    Buffer.BlockCopy(finalData, 0, partData, 0, finalData.Length);
                    Buffer.BlockCopy(prefixData, 0, partData, finalData.Length, prefixData.Length);
                    Buffer.BlockCopy(fileData, 0, partData, finalData.Length + prefixData.Length, fileData.Length);
                    finalData = partData;
                }

                return finalData;
            }
            
            return File.ReadAllBytes(Index[fileName].FullPath);
        }

        public string Extract(string fileName, string exportPath)
        {
            byte[] finalData = ExtractData(fileName);
            string finalPath;
            if (exportPath.Trim() == String.Empty)
            {
                finalPath = ArchiveInfo.DirectoryName + @"\" + fileName;
            }
            else
            {
                if (!Directory.Exists(exportPath.Trim()))
                {
                    throw new Exception("Selected export path does not exist.");
                }
                finalPath = exportPath.Trim() + @"\" + fileName;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(finalPath) ?? throw new InvalidOperationException());
            File.WriteAllBytes(finalPath, finalData);

            return ArchiveInfo.DirectoryName + @"\" + fileName;
        }

        public string SaveArchive(string archivePath)
        {
            if (archivePath.ToLower().EndsWith(".rpi"))
            {
                archivePath = Regex.Replace(archivePath, @"\.rpi$", ".rpa", RegexOptions.IgnoreCase);
            }
            
            if (!archivePath.ToLower().EndsWith(".rpa"))
            {
                archivePath += ".rpa";
            }

            string tmpPath = Regex.Replace(archivePath, @"\.rpa$", "", RegexOptions.IgnoreCase);
            tmpPath = tmpPath.Substring(0, Math.Min(100, tmpPath.Length - 1)) + "_" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + Guid.NewGuid().ToString("N");

            /*if (archivePath == _archivePath && _archivePath != String.Empty)
            {
                throw new Exception("Cannot overwrite same archive file that is loaded.");
            }*/

            string indexPath = Regex.Replace(archivePath, @"\.rpa$", ".rpi", RegexOptions.IgnoreCase);

            /*if (indexPath == _indexPath && _indexPath != String.Empty)
            {
                throw new Exception("Cannot overwrite same index file that is loaded.");
            }*/
            
            BuildArchive(archivePath, indexPath, tmpPath);

            return archivePath;
        }

        private void BuildArchive(string archivePath, string indexPath, string tmpPath)
        {
            try
            {
                if (!File.Exists(tmpPath + ".rpa"))
                {
                    File.WriteAllBytes(tmpPath + ".rpa", new byte[] { });
                }

                using (Stream stream = File.Open(tmpPath + ".rpa", FileMode.Truncate))
                {
                    int archiveOffset;
                    switch (ArchiveVersion)
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
                    foreach (KeyValuePair<string, ArchiveIndex> index in Index)
                    {
                        byte[] content = ExtractData(index.Key);

                        if (Padding > 0)
                        {
                            string paddingStr = String.Empty;
                            int paddingLength = rnd.Next(1, Padding);

                            while (paddingLength > 0)
                            {
                                paddingStr += Encoding.ASCII.GetString(new[] {(byte) rnd.Next(1, 255)});
                                paddingLength--;
                            }

                            byte[] paddingBytes = Encoding.ASCII.GetBytes(paddingStr);
                            archiveOffset += paddingBytes.Length;
                        }

                        stream.Position = archiveOffset;
                        stream.Write(content, 0, content.Length);

                        List<object[]> indexData = new List<object[]>();
                        if (CheckVersion(ArchiveVersion, Version.RPA_3) ||
                            CheckVersion(ArchiveVersion, Version.RPA_3_2))
                        {
                            indexData.Add(new object[]
                                {archiveOffset ^ Step, content.Length ^ Step, ""}); // Last is prefix
                        }
                        else
                        {
                            indexData.Add(new object[] {archiveOffset, content.Length});
                        }

                        archiveOffset += content.Length;

                        indexes.Add(index.Value.TreePath, indexData);
                    }

                    byte[] pickledIndexes;
                    using (Pickler pickler = new Pickler())
                    {
                        pickledIndexes = pickler.dumps(indexes);
                    }

                    byte[] fileCompressed = ZlibStream.CompressBuffer(pickledIndexes);

                    if (!CheckVersion(ArchiveVersion, Version.RPA_1))
                    {
                        stream.Position = archiveOffset;
                        stream.Write(fileCompressed, 0, fileCompressed.Length);

                        string headerContent = String.Empty;

                        switch (ArchiveVersion)
                        {
                            case Version.RPA_3_2:
                                headerContent = ArchiveMagic.RPA_3_2 + archiveOffset.ToString("x").PadLeft(16, '0') +
                                                " " +
                                                Step.ToString("x").PadLeft(8, '0') + "\n";
                                break;
                            case Version.RPA_3:
                                headerContent = ArchiveMagic.RPA_3 + archiveOffset.ToString("x").PadLeft(16, '0') +
                                                " " +
                                                Step.ToString("x").PadLeft(8, '0') + "\n";
                                break;
                            case Version.RPA_2:
                                headerContent = ArchiveMagic.RPA_2 + archiveOffset.ToString("x").PadLeft(16, '0') +
                                                "\n";
                                break;
                        }

                        byte[] headerContentByte = Encoding.UTF8.GetBytes(headerContent);

                        stream.Position = 0;
                        stream.Write(headerContentByte, 0, headerContentByte.Length);
                    }
                    else
                    {
                        File.WriteAllBytes(tmpPath + ".rpi", fileCompressed);
                    }
                }

                try
                {
                    // Test if archive is corrupted or not
                    RpaParser testParse = new RpaParser();
                    testParse.LoadArchive(tmpPath + ".rpa");
                }
                catch (Exception ex)
                {
                    throw new Exception("Validation of newly created archive failed. This usually means corrupted archive file after creation. No harm was done to original archive. Parser failed with following error during validation: " + ex.Message);
                }

                File.Copy(tmpPath + ".rpa", archivePath, true);
                File.Delete(tmpPath + ".rpa");
                if (File.Exists(tmpPath + ".rpi"))
                {
                    File.Copy(tmpPath + ".rpi", indexPath, true);
                    File.Delete(tmpPath + ".rpi");
                }
            }
            catch
            {
                if (File.Exists(tmpPath + ".rpa"))
                {
                    File.Delete(tmpPath + ".rpa");
                }
                if (File.Exists(tmpPath + ".rpi"))
                {
                    File.Delete(tmpPath + ".rpi");
                }

                throw;
            }
        }
    }

    // https://stackoverflow.com/a/28418846/3650856
    public class StructConverter
    {
        // We use this function to provide an easier way to make type-agnostic call via GetBytes method of the BitConverter class.
        // This means we can have much cleaner code below.
        private static byte[] TypeAgnosticGetBytes(object o)
        {
            switch (o)
            {
                case char c:
                    return BitConverter.GetBytes(c);
                case int i:
                    return BitConverter.GetBytes(i);
                case uint u:
                    return BitConverter.GetBytes(u);
                case long l:
                    return BitConverter.GetBytes(l);
                case ulong @ulong:
                    return BitConverter.GetBytes(@ulong);
                case short s:
                    return BitConverter.GetBytes(s);
                case ushort @ushort:
                    return BitConverter.GetBytes(@ushort);
                case byte _:
                case sbyte _:
                    return new[] { (byte)o };
                default:
                    throw new ArgumentException("Unsupported object type found");
            }
        }

        private static string GetFormatSpecifierFor(object o)
        {
            switch (o)
            {
                case char _:
                    return "c";
                case int _:
                    return "i";
                case uint _:
                    return "I";
                case long _:
                    return "q";
                case ulong _:
                    return "Q";
                case short _:
                    return "h";
                case ushort _:
                    return "H";
                case byte _:
                    return "B";
                case sbyte _:
                    return "b";
                default:
                    throw new ArgumentException("Unsupported object type found");
            }
        }

        /// <summary>
        /// Convert a byte array into an array of objects based on Python's "struct.unpack" protocol.
        /// </summary>
        /// <param name="fmt">A "struct.pack"-compatible format string</param>
        /// <param name="bytes">An array of bytes to convert to objects</param>
        /// <returns>Array of objects.</returns>
        /// <remarks>You are responsible for casting the objects in the array back to their proper types.</remarks>
        public static object[] Unpack(string fmt, byte[] bytes)
        {
            Debug.WriteLine("Format string is length {0}, {1} bytes provided.", fmt.Length, bytes.Length);

            // First we parse the format string to make sure it's proper.
            if (fmt.Length < 1) throw new ArgumentException("Format string cannot be empty.");

            bool endianFlip = false;
            if (fmt.Substring(0, 1) == "<")
            {
                Debug.WriteLine("  Endian marker found: little endian");
                // Little endian.
                // Do we need to flip endianness?
                if (BitConverter.IsLittleEndian == false) endianFlip = true;
                fmt = fmt.Substring(1);
            }
            else if (fmt.Substring(0, 1) == ">")
            {
                Debug.WriteLine("  Endian marker found: big endian");
                // Big endian.
                // Do we need to flip endianness?
                if (BitConverter.IsLittleEndian) endianFlip = true;
                fmt = fmt.Substring(1);
            }

            // Now, we find out how long the byte array needs to be
            int totalByteLength = 0;
            foreach (char c in fmt)
            {
                Debug.WriteLine("  Format character found: {0}", c);
                switch (c)
                {
                    case 'q':
                    case 'Q':
                        totalByteLength += 8;
                        break;
                    case 'i':
                    case 'I':
                        totalByteLength += 4;
                        break;
                    case 'h':
                    case 'H':
                        totalByteLength += 2;
                        break;
                    case 'b':
                    case 'B':
                    case 'x':
                        totalByteLength += 1;
                        break;
                    default:
                        throw new ArgumentException("Invalid character found in format string.");
                }
            }

            Debug.WriteLine("Endianness will {0}be flipped.", (object)(endianFlip ? "" : "NOT "));
            Debug.WriteLine("The byte array is expected to be {0} bytes long.", totalByteLength);

            // Test the byte array length to see if it contains as many bytes as is needed for the string.
            if (bytes.Length != totalByteLength) throw new ArgumentException("The number of bytes provided does not match the total length of the format string.");

            // Ok, we can go ahead and start parsing bytes!
            int byteArrayPosition = 0;
            var outputList = new List<object>();

            Debug.WriteLine("Processing byte array...");
            foreach (char c in fmt)
            {
                byte[] buf;
                switch (c)
                {
                    case 'q':
                        outputList.Add(BitConverter.ToInt64(bytes, byteArrayPosition));
                        byteArrayPosition += 8;
                        Debug.WriteLine("  Added signed 64-bit integer.");
                        break;
                    case 'Q':
                        outputList.Add(BitConverter.ToUInt64(bytes, byteArrayPosition));
                        byteArrayPosition += 8;
                        Debug.WriteLine("  Added unsigned 64-bit integer.");
                        break;
                    case 'i':
                        outputList.Add(BitConverter.ToInt32(bytes, byteArrayPosition));
                        byteArrayPosition += 4;
                        Debug.WriteLine("  Added signed 32-bit integer.");
                        break;
                    case 'I':
                        outputList.Add(BitConverter.ToUInt32(bytes, byteArrayPosition));
                        byteArrayPosition += 4;
                        Debug.WriteLine("  Added unsigned 32-bit integer.");
                        break;
                    case 'h':
                        outputList.Add(BitConverter.ToInt16(bytes, byteArrayPosition));
                        byteArrayPosition += 2;
                        Debug.WriteLine("  Added signed 16-bit integer.");
                        break;
                    case 'H':
                        if (endianFlip)
                        {
                            var deezBytes = bytes.Reverse().Skip(byteArrayPosition).Take(2).ToArray();
                            outputList.Add(BitConverter.ToUInt16(deezBytes, 0));
                        }
                        else
                        {
                            outputList.Add(BitConverter.ToUInt16(bytes, byteArrayPosition));
                        }

                        byteArrayPosition += 2;
                        Debug.WriteLine("  Added unsigned 16-bit integer.");
                        break;
                    case 'b':
                        buf = new byte[1];
                        Array.Copy(bytes, byteArrayPosition, buf, 0, 1);
                        outputList.Add((sbyte)buf[0]);
                        byteArrayPosition++;
                        Debug.WriteLine("  Added signed byte");
                        break;
                    case 'B':
                        buf = new byte[1];
                        Array.Copy(bytes, byteArrayPosition, buf, 0, 1);
                        outputList.Add(buf[0]);
                        byteArrayPosition++;
                        Debug.WriteLine("  Added unsigned byte");
                        break;
                    case 'x':
                        byteArrayPosition++;
                        Debug.WriteLine("  Ignoring a byte");
                        break;
                    default:
                        throw new ArgumentException("You should not be here.");
                }
            }
            return outputList.ToArray();
        }

        /// <summary>
        /// Convert an array of objects to a byte array, along with a string that can be used with Unpack.
        /// </summary>
        /// <param name="items">An object array of items to convert</param>
        /// <param name="littleEndian">Set to False if you want to use big endian output.</param>
        /// <param name="neededFormatStringToRecover">Variable to place an 'Unpack'-compatible format string into.</param>
        /// <returns>A Byte array containing the objects provided in binary format.</returns>
        public static byte[] Pack(object[] items, bool littleEndian, out string neededFormatStringToRecover)
        {

            // make a byte list to hold the bytes of output
            var outputBytes = new List<byte>();

            // should we be flipping bits for proper endianness?
            bool endianFlip = (littleEndian != BitConverter.IsLittleEndian);

            // start working on the output string
            string outString = (littleEndian == false ? ">" : "<");

            // convert each item in the objects to the representative bytes
            foreach (object o in items)
            {
                byte[] theseBytes = TypeAgnosticGetBytes(o);
                if (endianFlip) theseBytes = theseBytes.Reverse().ToArray();
                outString += GetFormatSpecifierFor(o);
                outputBytes.AddRange(theseBytes);
            }

            neededFormatStringToRecover = outString;

            return outputBytes.ToArray();

        }

        public static byte[] Pack(object[] items)
        {
            string dummy = "";
            return Pack(items, true, out dummy);
        }
    }
}