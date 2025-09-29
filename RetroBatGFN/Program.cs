using IWshRuntimeLibrary;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace RetroBatGFN
{
    internal class Program
    {
        static string globalMainDirectory = "";
        private static string currentPath = Directory.GetCurrentDirectory();

        // Import the FindWindow function from user32.dll
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Import the PostMessage function from user32.dll
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        const int WM_CLOSE = 0x0010;

        static async Task Main(string[] args)
        {
            await Startup();
            await RetroBatInstall();
            await MiscAppsInstall();
            await DesktopInstall();

            string retroBatExe = Path.Combine(globalMainDirectory, "RetroBat", "RetroBat.exe");
            Process.Start(retroBatExe);
        }

        static async Task Startup()
        {
            string asgardPath = @"C:\Asgard";
            string mainPathJson = "https://github.com/dpadGuy/RetroBatGFNThings/raw/refs/heads/main/directory.json";

            // GFN Enviorment check
            if (!Directory.Exists(asgardPath))
            {
                Console.WriteLine("[!] You are not running RetroBatGFN from a GFN enviorment, RetroBatGFN is gonna close...");

                Thread.Sleep(3000);
                Environment.Exit(0);
            }

            // Prepare the directory
            using (WebClient webClient = new WebClient())
            {
                string json = webClient.DownloadString(mainPathJson);

                var directory = JsonConvert.DeserializeObject<List<MainDirectory>>(json);

                if(Directory.Exists(globalMainDirectory))
                {
                    return; // Skip directory creation if it already exists
                }
                else
                {
                    Console.WriteLine("[+] Preparing the directory...");

                    foreach (var config in directory)
                    {
                        globalMainDirectory = config.directory;

                        Directory.CreateDirectory(config.directory);
                    }
                }
            }
        }

        static async Task RetroBatInstall()
        {
            string retroBatConfig = "https://github.com/dpadGuy/RetroBatGFNThings/raw/refs/heads/main/retrobat.json";
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string retroBatShortcut = Path.Combine(desktopPath, "RetroBat.lnk");
            string retroBatExe = Path.Combine(globalMainDirectory, "RetroBat", "RetroBat.exe");

            using (WebClient webClient = new WebClient())
            {
                string json = webClient.DownloadString(retroBatConfig);
                var retroConfig = JsonConvert.DeserializeObject<List<RetroBatConfig>>(json);
                var config = retroConfig[0];

                if (Directory.Exists(config.retroBatDir))
                {
                    WshShell shellB = new WshShell();
                    IWshShortcut shortcutB = (IWshShortcut)shellB.CreateShortcut(retroBatShortcut); // Make shortcut on desktop
                    shortcutB.TargetPath = retroBatExe;

                    shortcutB.Save();

                    return; // Skip RetroBat install if it already exists
                }

                await webClient.DownloadFileTaskAsync(new Uri(config.retroBatLink), $"{currentPath}\\RetroBat.exe");

                Console.WriteLine($"[!] You are gonna be prompted to install RetroBat, install using the folder below, DO NOT INSTALL COMPONENTS\n{globalMainDirectory}\\RetroBat");

                Process.Start($"{currentPath}\\RetroBat.exe").WaitForExit();

                WshShell shellA = new WshShell();
                IWshShortcut shortcutA = (IWshShortcut)shellA.CreateShortcut(retroBatShortcut); // Make shortcut on desktop
                shortcutA.TargetPath = retroBatExe;

                shortcutA.Save();
            }
        }

        static async Task MiscAppsInstall()
        {
            string jsonUrl = "https://github.com/dpadGuy/RetroBatGFNThings/raw/refs/heads/main/apps.json";

            using (WebClient webClient = new WebClient())
            {
                string json = await webClient.DownloadStringTaskAsync(jsonUrl);
                List<Apps> apps = JsonConvert.DeserializeObject<List<Apps>>(json);

                foreach (var app in apps)
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + $"\\{app.name}.lnk";
                    string appName = Path.Combine(globalMainDirectory, app.name);
                    string appExePath = Path.Combine(globalMainDirectory, app.exeName);
                    string appPath = Path.Combine(globalMainDirectory, app.name, app.exeName);

                    if (!Directory.Exists(appName))
                    {
                        if (app.fileExtension == "zip")
                        {
                            Console.WriteLine("[+] Installing " + app.name);

                            await webClient.DownloadFileTaskAsync(new Uri(app.url), $"{appName}.zip");

                            ZipFile.ExtractToDirectory($"{appName}.zip", appName);

                            System.IO.File.Delete($"{appName}.zip");

                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(desktopPath);
                            shortcut.TargetPath = appPath;
                            shortcut.WorkingDirectory = System.IO.Path.GetDirectoryName(appPath);

                            shortcut.Save();

                            if (app.run == "true")
                            {
                                Process.Start(appPath);
                            }
                        }

                        if (app.fileExtension == "exe")
                        {
                            Console.WriteLine("[+] Installing " + app.name);

                            await webClient.DownloadFileTaskAsync(new Uri(app.url), appExePath);

                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(desktopPath);
                            shortcut.TargetPath = appExePath;
                            shortcut.WorkingDirectory = System.IO.Path.GetDirectoryName(globalMainDirectory);

                            shortcut.Save();

                            if (app.run == "true")
                            {
                                Process.Start(appExePath);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("[!] " + app.name + " Already exists.");

                        if (app.fileExtension == "zip")
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(desktopPath);
                            shortcut.TargetPath = appPath;
                            shortcut.WorkingDirectory = System.IO.Path.GetDirectoryName(globalMainDirectory);

                            shortcut.Save();

                            if (app.run == "true")
                            {
                                Process.Start(appPath);
                            }
                        }

                        if (app.fileExtension == "exe")
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(desktopPath);
                            shortcut.TargetPath = appExePath;
                            shortcut.WorkingDirectory = System.IO.Path.GetDirectoryName(globalMainDirectory);

                            shortcut.Save();

                            if (app.run == "true")
                            {
                                Process.Start(appExePath);
                            }
                        }
                    }
                }
            }
        }
        static async Task DesktopInstall()
        {
            string jsonUrl = "https://github.com/dpadGuy/RetroBatGFNThings/raw/refs/heads/main/desktop.json";
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            IntPtr hWndSeelen = FindWindow(null, "CustomExplorer");

            // Check if the window handle is valid
            if (hWndSeelen != IntPtr.Zero)
            {
                PostMessage(hWndSeelen, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }

            using (WebClient webClient = new WebClient())
            {
                string json = await webClient.DownloadStringTaskAsync(jsonUrl);
                List<DesktopInfo> desktopInfo = JsonConvert.DeserializeObject<List<DesktopInfo>>(json);

                foreach (var desktops in desktopInfo)
                {
                    string appPath = Path.Combine(globalMainDirectory, desktops.name);
                    string zipFile = Path.Combine(globalMainDirectory, desktops.name + ".zip");
                    string exePath = Path.Combine(appPath, desktops.exeName);
                    string taskbarFixerPath = string.IsNullOrEmpty(desktops.taskbarFixer) ? "" : Path.Combine(appPath, desktops.taskbarFixer);
                    string roamingPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

                    if (!Directory.Exists(appPath))
                    {
                        await webClient.DownloadFileTaskAsync(new Uri(desktops.url), zipFile);

                        ZipFile.ExtractToDirectory(zipFile, appPath);

                        if (desktops.name == "WinXShell_x64")
                        {
                            Process.Start(exePath);
                            Process.Start(taskbarFixerPath);
                        }
                    }
                    else
                    {
                        Console.WriteLine("[!] " + desktops.name + " Already exists.");

                        if (desktops.name == "WinXShell_x64")
                        {
                            Process.Start(exePath);
                            Process.Start(taskbarFixerPath);
                        }
                    }
                }
            }
        }

        // Root myDeserializedClass = JsonConvert.DeserializeObject<List<Root>>(myJsonResponse);
        public class RetroBatConfig
        {
            public string retroBatLink { get; set; }
            public string directoryCreate { get; set; }
            public string retroBatDir { get; set; }
        }

        public class Apps
        {
            public string name { get; set; }
            public string fileExtension { get; set; }
            public string exeName { get; set; }
            public string run { get; set; }
            public string url { get; set; }
        }
        public class DesktopInfo
        {
            public string name { get; set; }
            public string exeName { get; set; }
            public string taskbarFixer { get; set; }
            public string zipConfig { get; set; }
            public string run { get; set; }
            public string url { get; set; }
        }

        public class MainDirectory
        {
            public string directory { get; set; }
        }
    }
}
