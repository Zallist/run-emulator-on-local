using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using RoboSharp;
using RoboSharp.Interfaces;

namespace EmulatorStarter
{
    internal static class Program
    {
        #region Configuration

        private static string EmulatorRemotePath =
            Path.GetFullPath(
                Environment.ExpandEnvironmentVariables(
                    System.Configuration.ConfigurationManager.AppSettings["EmulatorRemotePath"]));

        private static string EmulatorLocalPath =
            Path.GetFullPath(
                Environment.ExpandEnvironmentVariables(
                    System.Configuration.ConfigurationManager.AppSettings["EmulatorLocalPath"]));

        private static string EmulatorDesktopExe = System.Configuration.ConfigurationManager.AppSettings["EmulatorDesktopExe"];

        private static string EmulatorFullscreenExe = System.Configuration.ConfigurationManager.AppSettings["EmulatorFullscreenExe"];

        private const string LIBRARY_FILES_RELATIVE_PATH = "library\\files";

        #endregion

        private const string APPLICATION_ID = "5378a98e-8d56-4f3d-aa3c-c069e056dcc1";

        static async Task Main(string[] args)
        {
            System.Windows.Forms.Application.EnableVisualStyles();

            Mutex mutex;

            if (Mutex.TryOpenExisting(APPLICATION_ID, out _))
            {
                IsAlreadyRunning();
                return;
            }

            mutex = new Mutex(false, APPLICATION_ID, out var createdNew);

            if (!createdNew)
            {
                IsAlreadyRunning();
                return;
            }

            var argsList = new List<string>(args);

            var executable = EmulatorDesktopExe;
            if (argsList.Remove("fullscreen"))
                executable = EmulatorFullscreenExe;
            if (argsList.Remove("desktop"))
                executable = EmulatorDesktopExe;

            var updateAllRemoteFiles = argsList.Remove("update-all-remote-files");

            using (mutex)
            {
                EnsureInternetOptionsTrust();
                await CopyFromSourceToTarget(EmulatorRemotePath, EmulatorLocalPath);
                await EnsureLibraryFilesDirectoryJunction();

                await StartEmulator(executable, argsList);

                if (updateAllRemoteFiles)
                {
                    await CopyFromSourceToTarget(EmulatorLocalPath, EmulatorRemotePath);
                }
                else
                {
                    await CopySpecificsToRemote();
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
                cmd.StartInfo.FileName = Path.Combine(EmulatorLocalPath, exe);
                cmd.StartInfo.Arguments = string.Join(" ", args);

                cmd.StartInfo.UseShellExecute = false;

                cmd.StartInfo.WorkingDirectory = EmulatorLocalPath;

                await StartAndWaitAsync(cmd);
            }
        }

        private static async Task CopyFromSourceToTarget(string source, string target)
        {
            using (var robocmd = new RoboSharp.RoboCommand(source, target))
            {
                ConfigureRoboCommand(robocmd);

                robocmd.SelectionOptions.ExcludedDirectories.Add(Path.Combine(source, LIBRARY_FILES_RELATIVE_PATH));

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
            foreach (var copyAfter in System.Configuration.ConfigurationManager.AppSettings["CopyFoldersBackAfterRun"].Split(','))
            {
                var source = Path.Combine(EmulatorLocalPath, copyAfter);
                var target = Path.Combine(EmulatorRemotePath, copyAfter);

                using (var robocmd = new RoboSharp.RoboCommand(source, target))
                {
                    ConfigureRoboCommand(robocmd);
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

            robocmd.SelectionOptions.ExcludeOlder = true;
            robocmd.SelectionOptions.ExcludeJunctionPoints = true;
        }

        private static async Task EnsureLibraryFilesDirectoryJunction()
        {
            var libraryFilesDir = new DirectoryInfo(Path.Combine(EmulatorLocalPath, LIBRARY_FILES_RELATIVE_PATH));

            if (!libraryFilesDir.Exists || !libraryFilesDir.Attributes.HasFlag(FileAttributes.ReparsePoint))
            {
                if (libraryFilesDir.Exists)
                    libraryFilesDir.MoveTo(Path.Combine(EmulatorLocalPath, LIBRARY_FILES_RELATIVE_PATH + ".bak"));

                // make junction
                using (var cmd = new System.Diagnostics.Process())
                {
                    cmd.StartInfo.FileName = "cmd.exe";
                    cmd.StartInfo.UseShellExecute = true;
                    cmd.StartInfo.CreateNoWindow = true;
                    cmd.StartInfo.Verb = "runas";
                    cmd.StartInfo.Arguments = $"/c mklink /D \"{Path.Combine(EmulatorLocalPath, LIBRARY_FILES_RELATIVE_PATH)}\" \"{Path.Combine(EmulatorRemotePath, LIBRARY_FILES_RELATIVE_PATH)}\"";

                    await StartAndWaitAsync(cmd);
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

        private static void EnsureInternetOptionsTrust()
        {
            var remoteUNCPath = MappedDriveResolver.ResolveToRootUNC(EmulatorRemotePath);

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
