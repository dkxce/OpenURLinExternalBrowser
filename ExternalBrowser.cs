//
// C# 
// OpenUrl in ExternalBrowser
// v 0.2, 21.05.2024
// https://github.com/dkxce
// en,ru,1251,utf-8
//

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace dkxce.OpenUrl
{
    internal class ExternalBrowser
    {
        #pragma warning disable format // @formatter:off
        public enum OpenUrlMode
        {
            /// <summary>
            /// Code Default
            /// </summary>
            NotSet                    = 0000,
            /// <summary>
            ///  User Default
            /// </summary>
            UserDefault               = 0001,
            /// <summary>
            /// System Default
            /// </summary>
            SystemDefault             = 0002,
            /// <summary>
            /// Command Line Interface/Interpreter Deafult
            /// </summary>
            CLIDefault                = 0003,    
            /// <summary>
            /// PowerShell Default
            /// </summary>
            PowerShellDefault         = 0004,
            /// <summary>
            /// ShellAPI Default
            /// </summary>
            ShellAPIDefault           = 0005,
            /// <summary>
            /// Explorer Default
            /// </summary>
            ExplorerDefault           = 0006,
            /// <summary>
            /// url.dll,FileProtocolHandler
            /// </summary>
            FileProtocolHandler       = 0007,
            /// <summary>
            /// url.dll,OpenURL
            /// </summary>
            OpenURL                   = 0008,
            /// <summary>
            /// Preferal Browsers, see `PreferalBrowsers` dict
            /// </summary>
            PreferalBrowsers          = 0009,
            /// <summary>
            /// Registry Defaults (User and System)
            /// </summary>
            RegistryDefaults          = 0010,
            /// <summary>
            /// Shells Defaults (CLI, PowerShell, Shell)
            /// </summary>
            AllShellsDefauls          = 0011,
            /// <summary>
            /// Handlers Defaults (Explorer, url.dll)
            /// </summary>
            HandlersDefaults          = 0012,
        }
        #pragma warning restore format // @formatter:on

        #region WinAPI Methods

        [DllImport("psapi.dll")]
        private static extern uint GetProcessImageFileName(IntPtr hProcess, [Out] StringBuilder lpImageFileName, [In][MarshalAs(UnmanagedType.U4)] int nSize);

        [DllImport("psapi.dll")]
        private static extern uint GetModuleFileNameEx(IntPtr hProcess, IntPtr hModule, [Out] StringBuilder lpBaseName, [In][MarshalAs(UnmanagedType.U4)] int nSize);

        #endregion WinAPI Methods

        #region BROWSERS PRIORITY

        /// <summary>
        ///     Preferal External Browsers for Open Method.
        ///     {"process name with no extention",   "relative path slash-ending"}:
        ///     Firefox->Chrome->Brave->Yandex->Edge->Opera
        /// </summary>
        public static Dictionary<string, string> PreferalBrowsers = new Dictionary<string, string>() // In Priority
        {
            { "firefox",          @"Mozilla Firefox\" },                             // FireFox // %USERPROFILE%\Downloads
            { "chrome",           @"Google\Chrome\Application\" },                   // Chrome  // %USERPROFILE%\Downloads
            { "brave",            @"BraveSoftware\Brave-Browser\Application\" },     // Brave   // %USERPROFILE%\Downloads
            { "browser",          @"Yandex\YandexBrowser\Application\" },            // Yandex  // %USERPROFILE%\Downloads
         // { "safari",           @"Safari\" },                                      // Safari
            { "msedge",           @"Microsoft\Edge\Application\" },                  // Edge    // %USERPROFILE%\Downloads              
            { "opera",            @"Opera\" },                                       // Opera                
         // { "vivaldi",          @"Vivaldi\Application\" },                         // Vivaldi
         // { "atom",             @"Mail.Ru\Atom\Application\" },                    // Atom    // %USERPROFILE%\Downloads
         // { "maxthon",          @"Maxthon\Application\" },                         // Maxthon
         // { "spark",            @"baidu\Baidu Browser" },                          // Baidu Spark
         // { "k-meleon",         @"K-Meleon\" },                                    // K-Meleon
        };

        #endregion BROWSERS PRIORITY        

        /// <summary>
        ///     Synchronous Launch (wait result)
        /// </summary>
        public static bool SynchronousLaunch = false;

        #region Browsers

        /// <summary>
        ///     Get Executable Path of Browser
        /// </summary>
        /// <param name="exe">Executable File Name</param>
        /// <returns></returns>
        public static string GetCustomBrowser(string exe = "iexplore.exe")
        {
            if (string.IsNullOrEmpty(exe)) return null;
            RegistryKey regKey = null;
            try
            {
                regKey = Registry.ClassesRoot.OpenSubKey($"Applications\\{exe}\\DefaultIcon", false);
                string name = regKey.GetValue(null).ToString().ToLower().Replace("" + (char)34, "");
                if (!name.EndsWith("exe")) name = name.Substring(0, name.LastIndexOf(".exe") + 4);
                return name.Trim('"');
            }
            catch { } finally { if (regKey != null) regKey.Close(); };
            try
            {
                regKey = Registry.ClassesRoot.OpenSubKey($"Applications\\{exe}\\shell\\open\\command", false);
                string name = regKey.GetValue(null).ToString().ToLower().Replace("" + (char)34, "");
                if (!name.EndsWith("exe")) name = name.Substring(0, name.LastIndexOf(".exe") + 4);
                return name.Trim('"');
            }
            catch { } finally { if (regKey != null) regKey.Close(); };
            return null;
        }

        /// <summary>
        ///     Get User Default Browser
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public static string GetUserDefaultBrowser(out string args)
        {
            args = null;
            foreach (string proto in new string[] { "http", "https" })
            {
                args = null;
                RegistryKey regKey = null;
                string regVal = null;
                int exeLen = 4;
                try
                {
                    regKey = Registry.CurrentUser.OpenSubKey($"SOFTWARE\\Microsoft\\Windows\\Shell\\Associations\\URLAssociations\\{proto}\\UserChoice", false);
                    if (regKey == null) continue;
                    regVal = regKey.GetValue("ProgID", "").ToString();
                    regKey.Close();
                    if (string.IsNullOrEmpty(regVal)) continue;

                    regKey = Registry.ClassesRoot.OpenSubKey($"{regVal}\\shell\\open\\command");
                    if (regKey == null) continue;
                    regVal = regKey.GetValue(null, "").ToString();
                    regKey.Close();
                    if (string.IsNullOrEmpty(regVal)) continue;

                    int exepos = Math.Max(regVal.ToLower().IndexOf(".exe\""), regVal.ToLower().IndexOf(".exe\'"));
                    if (exepos > 0) exeLen++; else exepos = regVal.ToLower().IndexOf(".exe");
                    if (exepos > 0)
                    {
                        string fName = regVal.Substring(0, exepos + exeLen);
                        args = regVal.Replace(fName, "").Trim();
                        if (!args.Contains("%1")) args += " \"%1\"";
                        args = args.Trim();
                        return fName.Trim(new char[] { ' ', '\"', '\'' });
                    };
                }
                catch { } finally { if (regKey != null) regKey.Close(); };
            };
            return null;
        }

        /// <summary>
        ///     Get System Defauult Browser
        /// </summary>
        /// <returns></returns>
        public static string GetSystemDefaultBrowser()
        {
            string name = string.Empty;
            RegistryKey regKey = null;
            try
            {
                regKey = Registry.ClassesRoot.OpenSubKey("HTTP\\shell\\open\\command", false);
                name = regKey.GetValue(null).ToString().ToLower().Replace("" + (char)34, "");
                if (!name.EndsWith("exe")) name = name.Substring(0, name.LastIndexOf(".exe") + 4);
            }
            catch (Exception ex) { name = string.Format("ERROR: An exception of type: {0} occurred in method: {1}", ex.GetType(), ex.TargetSite); }
            finally { if (regKey != null) regKey.Close(); };
            return name;
        }

        #endregion Browsers

        #region Open

        /// <summary>
        ///     Open URI in External Browser
        /// </summary>
        /// <param name="url">URI</param>
        /// <param name="openMode">Open Mode</param>
        /// <returns></returns>
        public static bool Open(string url, OpenUrlMode openMode = OpenUrlMode.NotSet) // with searching //
        {
            Task<bool> t = new Task<bool>(() =>
            {                
                // Priority  1 //
                if (openMode == OpenUrlMode.NotSet || openMode == OpenUrlMode.RegistryDefaults || openMode == OpenUrlMode.UserDefault)
                {
                    try {
                        string defaultBrowser = GetUserDefaultBrowser(out string args);
                        if ((!string.IsNullOrEmpty(defaultBrowser)) && (!string.IsNullOrEmpty(args)))
                        {
                            args = args.Replace("'%1'", $"'{url}'").Replace("\"%1\"", $"\"{url}\"").Replace("%1", $"\"{url}\"");
                            Process.Start(new ProcessStartInfo() { FileName = defaultBrowser, Arguments = args });
                            return true;
                        };
                    } catch { };
                };

                // Priority  2 //
                if (openMode == OpenUrlMode.NotSet || openMode == OpenUrlMode.RegistryDefaults || openMode == OpenUrlMode.SystemDefault)
                {
                    try { Process.Start(GetSystemDefaultBrowser(), url); return true; } catch { };
                    try { Process.Start(new ProcessStartInfo() { FileName = GetCustomBrowser(), Arguments = $"\"{url}\"" }); return true; } catch { };
                };

                // Priority  3 //
                if (openMode == OpenUrlMode.NotSet || openMode == OpenUrlMode.AllShellsDefauls || openMode == OpenUrlMode.CLIDefault)
                {
                    // -- LOW DELAY
                    try { Process.Start(new ProcessStartInfo() { FileName = "cmd.exe", Arguments = $"/C start \"\" \"{url}\"", CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden }); return true; } catch { };
                };                               

                // Priority  4 // 
                if (openMode == OpenUrlMode.NotSet || openMode == OpenUrlMode.AllShellsDefauls || openMode == OpenUrlMode.PowerShellDefault)
                {
                    // -- BIG DELAY
                    try { Process.Start(new ProcessStartInfo() { FileName = "powershell.exe", Arguments = $" -command \"Start-Process '{url}'\"", CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden }); return true; } catch { };
                };

                // Priority  5 //
                if (openMode == OpenUrlMode.NotSet || openMode == OpenUrlMode.AllShellsDefauls || openMode == OpenUrlMode.ShellAPIDefault)
                {
                    try { Process.Start(new ProcessStartInfo() { FileName = url, UseShellExecute = true }); return true; } catch { };
                    try { Process.Start(new ProcessStartInfo() { FileName = url, UseShellExecute = false }); return true; } catch { };
                    try
                    {
                        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) { Process.Start(new ProcessStartInfo(url.Replace("&", "^&")) { UseShellExecute = true }); return true; }
                        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) { Process.Start("xdg-open", url); return true; }
                        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) { Process.Start("open", url); return true; };
                    }
                    catch { };
                };

                // Priority  6 //
                if (openMode == OpenUrlMode.NotSet || openMode == OpenUrlMode.HandlersDefaults || openMode == OpenUrlMode.ExplorerDefault)
                {
                    // -- MED DELAY
                    try { Process.Start(new ProcessStartInfo() { FileName = "explorer.exe", Arguments = $"\"{url}\"" }); return true; } catch { };
                };

                // Priority  7 //
                if (openMode == OpenUrlMode.NotSet || openMode == OpenUrlMode.HandlersDefaults || openMode == OpenUrlMode.FileProtocolHandler)
                {
                    ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", $"/C rundll32 url.dll,FileProtocolHandler \"{url}\"") { CreateNoWindow = true };
                    try { Process.Start(psi); return true; } catch { };
                };

                // Priority  8 //
                if (openMode == OpenUrlMode.NotSet || openMode == OpenUrlMode.HandlersDefaults || openMode == OpenUrlMode.OpenURL)
                {
                    ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", $"/C rundll32 url.dll,OpenURL \"{url}\"") { CreateNoWindow = true };
                    try { Process.Start(psi); return true; } catch { };
                };

                // Priority  9 //
                if (openMode == OpenUrlMode.NotSet || openMode == OpenUrlMode.PreferalBrowsers)
                {
                    string folderProgramFiles64 = Environment.ExpandEnvironmentVariables("%ProgramW6432%").Trim('\\');
                    string folderProgramFiles32 = Environment.ExpandEnvironmentVariables("%ProgramFiles(x86)%").Trim('\\');
                    string folderLocalAppData = Environment.ExpandEnvironmentVariables("%LocalAppData%").Trim('\\');
                    foreach (KeyValuePair<string, string> bro in PreferalBrowsers)
                    {
                        try { Process.Start(Process.GetProcessesByName($"{bro.Key}")?[0]?.MainModule?.FileName, url); return true; } catch { };
                        try { Process.Start(GetExecutableNameFromProcessName(bro.Key, bro.Key), url); return true; } catch { };
                        try { Process.Start(new ProcessStartInfo() { FileName = GetCustomBrowser($"{bro.Key}.exe"), Arguments = url }); return true; } catch { };
                        try { Process.Start(new ProcessStartInfo() { FileName = $"{folderProgramFiles64}\\{bro.Value}{bro.Key}.exe", Arguments = url }); return true; } catch { };
                        try { Process.Start(new ProcessStartInfo() { FileName = $"{folderProgramFiles32}\\{bro.Value}{bro.Key}.exe", Arguments = url }); return true; } catch { };
                        try { Process.Start(new ProcessStartInfo() { FileName = $"{folderLocalAppData}\\{bro.Value}{bro.Key}.exe", Arguments = url }); return true; } catch { };
                    };
                };                                                
                
                // Errors      //
                return false;
            });
            t.Start();
            if (!SynchronousLaunch) return true;
            t.Wait();
            return t.Result;
        }

        /// <summary>
        ///     Open with Preferal Browsers, see `PreferalBrowsers` dict
        /// </summary>
        /// <param name="url">URI</param>
        /// <returns></returns>
        public static bool OpenWithPreferalBrowser(string url) => Open(url, OpenUrlMode.PreferalBrowsers);

        /// <summary>
        ///     Open with `url.dll,FileProtocolHandler`
        /// </summary>
        /// <param name="url">URI</param>
        /// <returns></returns>
        public static bool OpenWithProtocolHandler(string url) => Open(url, OpenUrlMode.FileProtocolHandler);
        
        /// <summary>
        ///     Open with `explorer.exe` or `url.dll`
        /// </summary>
        /// <param name="url">URI</param>
        /// <returns></returns>
        public static bool OpenWithProtocolHandlers(string url) => Open(url, OpenUrlMode.HandlersDefaults);

        /// <summary>
        ///     Open with CLI/PowerShell/Shell
        /// </summary>
        /// <param name="url">URI</param>
        /// <returns></returns>
        public static bool OpenWithShell(string url) => Open(url, OpenUrlMode.AllShellsDefauls);

        /// <summary>
        ///     Open with `url.dll`
        /// </summary>
        /// <param name="url">URI</param>
        /// <returns></returns>
        public static bool OpenUrl(string url) => Open(url, OpenUrlMode.OpenURL);

        #endregion Open

        #region PRIVATES

        private static string GetExecutableNameFromProcessName(string procName, string defvalue = null, bool tryExe = true)
        {
            try
            {
                IntPtr? h = Process.GetProcessesByName(procName)?[0]?.Handle;
                if ((!h.HasValue) && tryExe) h = Process.GetProcessesByName($"{procName}.exe")?[0]?.Handle;
                if (h.HasValue)
                {
                    StringBuilder sb = new StringBuilder(ushort.MaxValue);
                    GetModuleFileNameEx(h.Value, IntPtr.Zero, sb, sb.Capacity);
                    string res = sb.ToString();
                    if (!string.IsNullOrEmpty(res)) return res;
                };
            }
            catch { };
            return defvalue;
        }

        #endregion PRIVATES
    }

}
