using System;
using System.Collections.Generic;
using System.IO;

namespace RPA_Explorer
{
    public class Settings
    {
        private bool _loadingPending = false;
        private readonly string _settingsPath;
        
        public Settings(string settingsPath)
        {
            _settingsPath = settingsPath;
            LoadSettings();
        }

        private void StoreSettings()
        {
            if (!_loadingPending)
            {
                List<string> cfg = new List<string>();
                
                cfg.Add("language=" + GetLang().Name);
                if (!string.IsNullOrEmpty(GetPython()))
                {
                    cfg.Add("python=" + GetPython());
                }
                if (!string.IsNullOrEmpty(GetUnrpyc()))
                {
                    cfg.Add("unrpyc=" + GetUnrpyc());
                }
                if (!string.IsNullOrEmpty(GetArchive()))
                {
                    cfg.Add("archive=" + GetArchive());
                }

                if (File.Exists(_settingsPath))
                {
                    File.Delete(_settingsPath);
                }
                File.WriteAllLines(_settingsPath, cfg);
                
                LoadSettings();
            }
        }

        private void LoadSettings()
        {
            _loadingPending = true;
            
            if (File.Exists(_settingsPath))
            {
                string[] lines = File.ReadAllLines(_settingsPath);
                foreach (string line in lines)
                {
                    string cfg = String.Empty;
                    if (line.Contains("#"))
                    {
                        cfg = line.Split('#')[0].Trim();
                    }
                    else
                    {
                        cfg = line.Trim();
                    }

                    if (!string.IsNullOrEmpty(cfg))
                    {
                        if (cfg.Contains("="))
                        {
                            string[] cfgSplit = cfg.Split('=');
                            string name = cfgSplit[0].Trim().ToLower();
                            string value = String.Empty;
                            if (cfgSplit.Length > 1)
                            {
                                value = cfgSplit[1].Trim();
                            }

                            switch (name)
                            {
                                case "language":
                                    SetLang(value);
                                    break;
                                case "python":
                                    SetPython(value);
                                    break;
                                case "unrpyc":
                                    SetUnrpyc(value);
                                    break;
                                case "archive":
                                    SetArchive(value);
                                    break;
                            }
                        }
                    }
                }
            } 
            // Else use defaults

            _loadingPending = false;
        }
        
        /*
         *
         * LANGUAGE
         * 
         */

        public class Language
        {
            public string Name;
            public string Abbrev;

            public Language(string name = "English", string abbrev = "EN")
            {
                Name = name;
                Abbrev = abbrev;
            }
        }
        
        // Set language values and add to langList to enable language
        private static readonly Language English = new Language("English", "EN");
        private static readonly Language Test = new Language("TEST", "TST");

        // List of enabled languages
        public readonly Language[] LangList = {
            English
            /* *-/
            // For testing only
            , Test
            /* */
        };
        
        private Language _language = new Language("English", "EN");
        
        public Language GetLang()
        {
            return _language;
        }

        public void SetLang(string language)
        {
            bool isValid = false;
            foreach (Language lang in LangList)
            {
                if (language == lang.Name || language == lang.Abbrev)
                {
                    _language = lang;
                    isValid = true;
                    break;
                }
            }

            if (!isValid)
            {
                // Default
                _language = English;
            }

            StoreSettings();
        }
        
        /*
         *
         * PYTHON PATH
         * 
         */
        
        private string _python = String.Empty;
        
        public string GetPython()
        {
            return _python;
        }

        public void SetPython(string path)
        {
            if (File.Exists(path))
            {
                _python = path;
            }
            else
            {
                // Default
                _python = String.Empty;
            }

            StoreSettings();
        }
        
        /*
         *
         * UNRPYC PATH
         * 
         */
        
        private string _unrpyc = String.Empty;
        
        public string GetUnrpyc()
        {
            return _unrpyc;
        }

        public void SetUnrpyc(string path)
        {
            if (File.Exists(path))
            {
                _unrpyc = path;
            }
            else
            {
                // Default
                _unrpyc = String.Empty;
            }

            StoreSettings();
        }
        
        /*
         *
         * LAST ARCHIVE PATH
         * 
         */
        
        private string _archive = String.Empty;
        
        public string GetArchive()
        {
            return _archive;
        }

        public void SetArchive(string path)
        {
            if (File.Exists(path))
            {
                _archive = path;
            }
            else
            {
                // Default
                _unrpyc = String.Empty;
            }

            StoreSettings();
        }
    }
}