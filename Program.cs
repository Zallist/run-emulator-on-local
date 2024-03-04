using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
using RoboSharp;
using RoboSharp.Interfaces;

namespace EmulatorStarter
{
    public static class Config
    {
        public static readonly string EmulatorRemotePath =
            Path.GetFullPath(
                Environment.ExpandEnvironmentVariables(
                    System.Configuration.ConfigurationManager.AppSettings["EmulatorRemotePath"]));

        public static readonly string EmulatorLocalPath =
            Path.GetFullPath(
                Environment.ExpandEnvironmentVariables(
                    System.Configuration.ConfigurationManager.AppSettings["EmulatorLocalPath"]));

        public static readonly string EmulatorDesktopExe = System.Configuration.ConfigurationManager.AppSettings["EmulatorDesktopExe"];

        public static readonly string EmulatorFullscreenExe = System.Configuration.ConfigurationManager.AppSettings["EmulatorFullscreenExe"];

        public static readonly string[] CopyFoldersBackAfterRun = 
            System.Configuration.ConfigurationManager.AppSettings["CopyFoldersBackAfterRun"].Split(',') ??
                new string[0];

        public static readonly string[] EnsureProgramsRunningWithEmulator = 
            System.Configuration.ConfigurationManager.AppSettings["EnsureProgramsRunningWithEmulator"]?.Split(',') ??
                new string[0];

        public static readonly bool DisableXboxGuideButton;

        public const string LIBRARY_FILES_RELATIVE_PATH = "library\\files";

        public const string APPLICATION_ID = "5378a98e-8d56-4f3d-aa3c-c069e056dcc1";

        static Config()
        {
            if (!bool.TryParse(System.Configuration.ConfigurationManager.AppSettings["DisableXboxGuideButton"], out DisableXboxGuideButton))
                DisableXboxGuideButton = false;
        }
    }

    internal static class Program
    {

        static async Task Main(string[] args)
        {
            System.Windows.Forms.Application.EnableVisualStyles();

            Mutex mutex;

            if (Mutex.TryOpenExisting(Config.APPLICATION_ID, out _))
            {
                IsAlreadyRunning();
                return;
            }

            mutex = new Mutex(false, Config.APPLICATION_ID, out var createdNew);

            if (!createdNew)
            {
                IsAlreadyRunning();
                return;
            }

            var argsList = new List<string>(args);

            var executable = Config.EmulatorDesktopExe;
            if (argsList.Remove("fullscreen"))
                executable = Config.EmulatorFullscreenExe;
            if (argsList.Remove("desktop"))
                executable = Config.EmulatorDesktopExe;

            var updateAllRemoteFiles = argsList.Remove("update-all-remote-files");

            using (mutex)
            {
                using (var systemOptionApplier = new SystemOptionApplier())
                {
                    await CopyFromSourceToTarget(Config.EmulatorRemotePath, Config.EmulatorLocalPath, CopyActionFlags.Purge);
                    await EnsureLibraryFilesDirectoryJunction();

                    EnsureProgramsRunningWithEmulator();

                    await StartEmulator(executable, argsList);

                    if (updateAllRemoteFiles)
                    {
                        await CopyFromSourceToTarget(Config.EmulatorLocalPath, Config.EmulatorRemotePath, CopyActionFlags.Purge);
                    }
                    else
                    {
                        await CopySpecificsToRemote();
                    }
                }
            }
        }

        static void IsAlreadyRunning()
        {
            System.Windows.Forms.MessageBox.Show(
                "Already running",
                "Already running");

            Environment.Exit(128);
        }

        private static async Task StartEmulator(string exe, IEnumerable<string> args)
        {
            using (var cmd = new Process())
            {
                cmd.StartInfo.FileName = Path.Combine(Config.EmulatorLocalPath, exe);
                cmd.StartInfo.Arguments = string.Join(" ", args);

                cmd.StartInfo.UseShellExecute = false;

                cmd.StartInfo.WorkingDirectory = Config.EmulatorLocalPath;

                await StartAndWaitAsync(cmd);
            }
        }

        private static async Task CopyFromSourceToTarget(string source, string target, CopyActionFlags copyActionFlags = CopyActionFlags.Default)
        {
            using (var robocmd = new RoboSharp.RoboCommand(source, target, copyActionFlags))
            {
                ConfigureRoboCommand(robocmd);

                robocmd.SelectionOptions.ExcludedDirectories.Add(Path.Combine(source, Config.LIBRARY_FILES_RELATIVE_PATH));

                var task = robocmd.Start();

                await Task.WhenAny(
                    task,
                    Task.Delay(500));
                
                if (!task.IsCompleted)
                {
                    var dialogShownTCS = new TaskCompletionSource<bool>();
                    var thread = new Thread(() =>
                    {
                        try
                        {
                            // Let user know we're waiting
                            using (var copyingDialog = new CopyingDialog())
                            {
                                copyingDialog.TopMost = true;
                                copyingDialog.UseWaitCursor = true;

                                copyingDialog.Show();
                                dialogShownTCS.TrySetResult(true);
                                System.Windows.Forms.Application.Run(copyingDialog);
                            }
                        }
                        catch(Exception) { }
                    })
                    {  IsBackground = true };

                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();

                    await dialogShownTCS.Task;
                    await task;

                    thread.Abort();
                    thread.Join();
                }
            }
        }

        private static async Task CopySpecificsToRemote()
        {
            foreach (var copyAfter in Config.CopyFoldersBackAfterRun)
            {
                var source = Path.Combine(Config.EmulatorLocalPath, copyAfter);
                var target = Path.Combine(Config.EmulatorRemotePath, copyAfter);

                using (var robocmd = new RoboSharp.RoboCommand(source, target))
                {
                    ConfigureRoboCommand(robocmd);

                    robocmd.SelectionOptions.ExcludeOlder = true;

                    await robocmd.Start();
                }
            }
        }

        private static void ConfigureRoboCommand(RoboCommand robocmd)
        {
            robocmd.CopyOptions.MultiThreadedCopiesCount = Environment.ProcessorCount;
            robocmd.CopyOptions.UseUnbufferedIo = true;
            robocmd.CopyOptions.CopySubdirectoriesIncludingEmpty = true;
            robocmd.CopyOptions.Compress = true;
            robocmd.CopyOptions.CopySymbolicLink = false;

            robocmd.RetryOptions.RetryCount = 3;
            robocmd.RetryOptions.RetryWaitTime = 1;

            //robocmd.SelectionOptions.ExcludeOlder = true;
            robocmd.SelectionOptions.ExcludeJunctionPoints = true;
        }

        private static async Task EnsureLibraryFilesDirectoryJunction()
        {
            var libraryFilesDir = new DirectoryInfo(Path.Combine(Config.EmulatorLocalPath, Config.LIBRARY_FILES_RELATIVE_PATH));

            if (!libraryFilesDir.Exists || !libraryFilesDir.Attributes.HasFlag(FileAttributes.ReparsePoint))
            {
                if (libraryFilesDir.Exists)
                    libraryFilesDir.MoveTo(Path.Combine(Config.EmulatorLocalPath, Config.LIBRARY_FILES_RELATIVE_PATH + ".bak"));

                // make junction
                using (var cmd = new System.Diagnostics.Process())
                {
                    cmd.StartInfo.FileName = "cmd.exe";
                    cmd.StartInfo.UseShellExecute = true;
                    cmd.StartInfo.CreateNoWindow = true;
                    cmd.StartInfo.Verb = "runas";
                    cmd.StartInfo.Arguments = $"/c mklink /D " +
                        $"\"{Path.Combine(Config.EmulatorLocalPath, Config.LIBRARY_FILES_RELATIVE_PATH)}\" " +
                        $"\"{Path.Combine(Config.EmulatorRemotePath, Config.LIBRARY_FILES_RELATIVE_PATH)}\"";

                    await StartAndWaitAsync(cmd);
                }
            }
        }

        private static void EnsureProgramsRunningWithEmulator()
        {
            foreach (var ensureRunning in Config.EnsureProgramsRunningWithEmulator)
            {
                var exeName = Path.GetFileName(ensureRunning);

                if (Process.GetProcessesByName(exeName).Length == 0)
                {
                    using (var cmd = new System.Diagnostics.Process())
                    { 
                        cmd.StartInfo.FileName = ensureRunning;
                        cmd.StartInfo.WorkingDirectory = Path.GetDirectoryName(ensureRunning);
                        cmd.StartInfo.UseShellExecute = true;

                        cmd.Start();

                        ChildProcessTracker.AddProcess(cmd);
                    }
                }
            }
        }

        private static async Task<int> StartAndWaitAsync(Process cmd, bool track = true)
        {
            TaskCompletionSource<int> tcs = new TaskCompletionSource<int>();

            cmd.EnableRaisingEvents = true;

            cmd.Exited += (sender, args) =>
            {
                cmd.WaitForExit();
                tcs.TrySetResult(cmd.ExitCode);
            };

            cmd.Start();

            if (track)
                ChildProcessTracker.AddProcess(cmd);

            return await tcs.Task;
        }
    }

    internal class SystemOptionApplier : IDisposable
    {
        private readonly List<Tuple<RegistryKey, string, object, RegistryValueKind>> currentUserRegistryKeysToRevertTo = new List<Tuple<RegistryKey, string, object, RegistryValueKind>>();

        public SystemOptionApplier()
        {
            EnsureInternetOptionsTrust();

            if (Config.DisableXboxGuideButton)
            {
                DisableXboxGuideButton();
            }
        }

        public void Dispose()
        {
            foreach (var tuple in currentUserRegistryKeysToRevertTo)
            {
                var subkey = tuple.Item1;
                var valueName = tuple.Item2;
                var valueData = tuple.Item3;
                var valueKind = tuple.Item4;

                if (valueData == null)
                {
                    subkey.DeleteValue(valueName, false);
                }
                else
                {
                    subkey.SetValue(valueName, valueData, valueKind);
                }
            }

            foreach (var key in currentUserRegistryKeysToRevertTo.Select(x => x.Item1).Distinct())
                key.Dispose();

            currentUserRegistryKeysToRevertTo.Clear();
        }

        private void DisableXboxGuideButton()
        {
            var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\GameBar", true);

            if (key != null)
            {
                currentUserRegistryKeysToRevertTo.Add(Tuple.Create(
                    key,
                    "UseNexusForGameBarEnabled", 
                    key.GetValue("UseNexusForGameBarEnabled", null), 
                    key.GetValue("UseNexusForGameBarEnabled", null) != null ? key.GetValueKind("UseNexusForGameBarEnabled") : RegistryValueKind.None));

                key.SetValue("UseNexusForGameBarEnabled", 0, RegistryValueKind.DWord);
            }

            try
            {
                foreach (var gameBarProcess in Process.GetProcessesByName("GameBar.exe"))
                    gameBarProcess.Kill();
            }
            catch (Exception)
            {

            }
        }

        private void EnsureInternetOptionsTrust()
        {
            var remoteUNCPath = MappedDriveResolver.ResolveToRootUNC(Config.EmulatorRemotePath);

            if (!remoteUNCPath.StartsWith(@"\\"))
                return;

            var remoteIPAddressMatch = Regex.Match(remoteUNCPath, @"^\\\\(.+?)\\.+$");

            if (!remoteIPAddressMatch.Success)
                return;

            var remoteIPAddress = remoteIPAddressMatch.Groups[1].Value;

            // Ensure we have a registry key at HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\RangeX 
            // with a :Range that contains remoteIPAddress

            // This makes sure we trust the exes
            for (int i = 1; i < 100; i++)
            {
                var currentValue = Microsoft.Win32.Registry.GetValue(
                    @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range" + i,
                    @":Range",
                    null);

                if (currentValue == null)
                {
                    Microsoft.Win32.Registry.SetValue(
                        @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range" + i,
                        @"*",
                        1,
                        Microsoft.Win32.RegistryValueKind.DWord);

                    Microsoft.Win32.Registry.SetValue(
                        @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range" + i,
                        @":Range",
                        remoteIPAddress,
                        Microsoft.Win32.RegistryValueKind.String);

                    break;
                }
                else if (StringComparer.OrdinalIgnoreCase.Equals(currentValue, remoteIPAddress))
                {
                    break;
                }
            }

            // Ensure we don't prompt when opening exes
            if (!int.TryParse((Microsoft.Win32.Registry.GetValue(
                @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1",
                @"1806",
                null) ?? "").ToString(), out var v) || v != 0)
            {
                Microsoft.Win32.Registry.SetValue(
                    @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1",
                    @"1806",
                    0,
                    Microsoft.Win32.RegistryValueKind.DWord);
            }
        }
    }
}
