using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Ionic.Zlib;
using Razorvine.Pickle;

namespace RPA_Explorer
{
    public class RpaParser
    {
        // Inspiration: https://github.com/Shizmob/rpatool/blob/master/rpatool
        
        private const string RPA2_MAGIC = "RPA-2.0 ";
        private const string RPA3_MAGIC = "RPA-3.0 ";
        private const string RPA3_2_MAGIC = "RPA-3.2 ";

        private readonly string _filePath;
        private readonly string _firstLine;
        public readonly FileInfo _fileInfo;
        private readonly double _version;
        private readonly string[] _metadata;
        private readonly long _offset;
        private readonly long _step;
        private readonly SortedDictionary<string,ArchiveIndex> _indexes;

        public class ArchiveIndex {
            public long offset = 0;
            public long length = 0;
            public string prefix = String.Empty;
        }

        internal RpaParser(string filePath)
        {
            _filePath = filePath;
            _fileInfo = GetFileInfo();
            _firstLine = GetFirstLine();
            _version = GetVersion();
            
            if (_version == 2 || _version == 3 || _version == 3.2)
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

                if (firstLine == null)
                {
                    throw new Exception("File is either not valid RenPy Archive or version is not supported.");
                }

                return firstLine;
            }
        }

        private double GetVersion()
        {
            if (_firstLine.StartsWith(RPA3_2_MAGIC))
            {
                return 3.2;
            }

            if (_firstLine.StartsWith(RPA3_MAGIC))
            {
                return 3;
            }

            if (_firstLine.StartsWith(RPA2_MAGIC))
            {
                return 2;
            }

            if (_filePath.EndsWith(".rpi"))
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
            
            if (_version == 3)
            {
                for(var i = 2; i < _metadata.Length; i++)
                {
                    step ^= Convert.ToInt64(_metadata[i], 16);
                }
            }
            else if (_version == 3.2)
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
            
            using (BinaryReader reader = new BinaryReader(File.Open(_filePath, FileMode.Open), Encoding.UTF8))
            {
                if (_version == 2 || _version == 3 || _version == 3.2)
                {
                    reader.BaseStream.Seek(_offset, SeekOrigin.Begin);
                }

                byte[] fileCompressed = reader.ReadBytes((int) reader.BaseStream.Length);
                byte[] fileUncompressed = ZlibStream.UncompressBuffer(fileCompressed);
                using (Unpickler unpickler = new Unpickler())
                {
                    unpickledIndexes = unpickler.loads(fileUncompressed);
                }
            }
            
            // Standardize output
            foreach (DictionaryEntry kvp in (Hashtable) unpickledIndexes)
            {
                string key = kvp.Key as string;
                object[] value = (kvp.Value as ArrayList).ToArray()[0] as object[];
                ArchiveIndex index = new ArchiveIndex();
                if ((long) value.Length == 2)
                {
                    index.offset = (long) value.GetValue(0);
                    index.length = Convert.ToInt64(value.GetValue(1));
                }
                else
                {
                    index.offset = (long) value.GetValue(0);
                    index.length = Convert.ToInt64(value.GetValue(1));
                    index.prefix = (string) value.GetValue(2);
                }
                indexes.Add(key, index);
            }
            
            // Deobfuscate index data
            if (_version >= 3)
            {
                foreach (KeyValuePair<string,ArchiveIndex> kvp in indexes)
                {
                    indexes[kvp.Key].offset ^= _step;
                    indexes[kvp.Key].length ^= _step;
                }
            }

            return indexes;
        }

        public SortedDictionary<string, ArchiveIndex> GetFileList()
        {
            return _indexes;
        }

        public FileInfo GetArchiveInfo()
        {
            return _fileInfo;
        }

        public string Extract(string fileName, string exportPath)
        {
            if (!_indexes.ContainsKey(fileName))
            {
                throw new Exception("Specified file does not exist in RenPy Archive.");
            }
            
            using (BinaryReader reader = new BinaryReader(File.Open(_filePath, FileMode.Open), Encoding.UTF8))
            {
                reader.BaseStream.Seek(_indexes[fileName].offset, SeekOrigin.Begin);
                byte[] prefixData = Encoding.UTF8.GetBytes(_indexes[fileName].prefix);
                byte[] fileData = reader.ReadBytes((int) _indexes[fileName].length -  _indexes[fileName].prefix.Length);
                byte[] finalData = new byte[prefixData.Length + fileData.Length];
                Buffer.BlockCopy(prefixData, 0, finalData, 0, prefixData.Length);
                Buffer.BlockCopy(fileData, 0, finalData, prefixData.Length, fileData.Length);
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
        }
    }
}