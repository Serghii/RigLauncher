using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using RigLauncher;
using File = Google.Apis.Drive.v3.Data.File;

namespace GDrive
{
    class Launcher
    {
        private static Action unzipCompleteAction = () => { };
        private static Action StartCompleteAction = () => { };
        private static Action StopCurRigCompleteAction = () => { };
        private static Action DeleteOldFolderCompleteAction = () => { };
        private static Action ChangeVersionOnGSheetCompleteAction = () => { };

        private static readonly string Rig = "Rig";
        private static readonly string curDir = Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetEntryAssembly().Location, @"..\"));
        private static string[] scope = {DriveService.Scope.Drive};
        private static string applicationName = "GDrive";
        private static string pathZipNewVersion = "";
        private static string pathZipOldVersion = "";
        private static string pathFolderOldVersion = "";
        private static string pathFolderNewVersion = "";
        private static RigVersion Ver;
        private static ProcessStartInfo startInfo = new ProcessStartInfo();
        static void Main(string[] args)
        {
            Console.WriteLine("Launcher: v2.1");
            StopDuplicateLauncher();

            Ver = GSheet.ReadGSheet();
            if (Ver == null)
            {
                RigEx.WriteLineColors("Cannot Read Google Sheet".AddTimeStamp(), ConsoleColor.Red);
                Console.Read();
                DelayAndQuit();
            }

            Console.WriteLine($"curent version:{Ver.curVersion} new version:{Ver.newVersion}");
            pathFolderOldVersion = $"{curDir}{Ver.curVersion}";
            pathFolderNewVersion = $"{curDir}{Ver.newVersion}";
            pathZipOldVersion = $"{curDir}{Ver.curVersion}.zip";
            pathZipNewVersion = $"{curDir}{Ver.newVersion}.zip";
           if (Ver.newVersion <= Ver.curVersion)
            {
                Console.WriteLine("Update not need ");
                ReopenRig();
                DelayAndQuit();
            }

            GDrive.DownloadCompletedAction += DownloadCompleted;
            unzipCompleteAction += StopRigProgram;
            StopCurRigCompleteAction += ChangeVersionOnGSheet;
            ChangeVersionOnGSheetCompleteAction += StartProgram;
            StartCompleteAction += DeleteOldFiles;
           Console.WriteLine("Begin Update");
            GDrive.DownloadFileFromGDrive(Ver.ZipId, pathZipNewVersion);
            Console.WriteLine("end");
            DelayAndQuit();
        }

        private static void ChangeVersionOnGSheet()
        {
            GSheet.SendData(Ver.newVersion.ToString());
            ChangeVersionOnGSheetCompleteAction();
        }

        private static void StopDuplicateLauncher()
        {
            Process process = Process.GetCurrentProcess();
            var dupl = Process.GetProcessesByName(process.ProcessName);
            if (dupl.Length > 1 )
            {
                foreach (var p in dupl)
                {
                    if (p.Id != process.Id)
                        p.Kill();
                }
            }
        }

        private static void ReopenRig()
        {
            Process[] process = Process.GetProcessesByName(Rig);
            if (process != null || process.Length != 0)
            {
                for (int i = 0; i < process.Length; i++)
                {
                    process[i].CloseMainWindow();
                }
            }
            StartProgram();
            StartCompleteAction();
        }

        private static void StopRigProgram()
        {
            Process[] process = Process.GetProcessesByName(Rig);
            for (int i = 0; i < process.Length; i++)
            {
                process[i].CloseMainWindow();
            }
            StopCurRigCompleteAction();
        }
        private static  void StartProgram()
        {
            try
            {
                Console.WriteLine("start Rig ");
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = true;
                startInfo.FileName = $"{pathFolderNewVersion}\\{Rig}.exe";
                startInfo.WindowStyle = ProcessWindowStyle.Normal;
                Process.Start(startInfo);
            }
            catch (Exception e)
            {
                RigEx.WriteLineColors($"Start Program Error: {e.Message}",ConsoleColor.Red);
                Console.ReadLine();
            }
        }
        private static void DownloadCompleted()
        {
            StartUnzip();
        }

       private static  void StartUnzip()
        {
            try
            {
              Console.WriteLine("start UnZip");
              ZipFile.ExtractToDirectory(pathZipNewVersion, $"{curDir}\\{Ver.newVersion}");
              Console.WriteLine("UnZipDir - DONE");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Unzip Error: {e.Message}");
                if (!e.Message.Contains("already exists"))
                {
                    DelayAndQuit();
                }
            }
            unzipCompleteAction();
        }

        private static void DeleteOldFiles()
        {
            if (pathFolderOldVersion == pathFolderNewVersion)
                return;

            try
            {
                Directory.Delete(pathFolderOldVersion, true);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Delete Old Folder ERROR: {e.Message}");
            }

            try
            {
                if (System.IO.File.Exists(pathZipOldVersion))
                    System.IO.File.Delete(pathZipOldVersion);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Delete Old Zip ERROR: {e.Message}");
            }
            DeleteOldFolderCompleteAction();
        }
        private void ZipDirToParent(string dirPath)
        {
            string parent = Path.GetDirectoryName(dirPath);
            string name = Path.GetFileName(dirPath);
            string fileName = Path.Combine(parent, name + ".zip");
            ZipFile.CreateFromDirectory(dirPath,fileName);
        }
        private static void UnZipDir(string zipPath)
        {
            ZipFile.ExtractToDirectory(zipPath, @"D:\Mining\TMP");
        }
        private static void Extract(FileInfo FileToDecompress)
        {
            using (FileStream origStream = FileToDecompress.OpenRead())
            {
                string currentFileName = FileToDecompress.FullName;
                string newFileName = currentFileName.Remove(currentFileName.Length - FileToDecompress.Extension.Length);
            }
        }
        private static void UnZipFile(string zipPath, string filePath)
        {
            using (FileStream inputStream = new FileStream(zipPath, FileMode.OpenOrCreate,FileAccess.ReadWrite ) )
            {
                using (FileStream outputStream = new FileStream(filePath, FileMode.OpenOrCreate,FileAccess.ReadWrite))
                {
                    using (GZipStream gzip = new GZipStream(inputStream,CompressionMode.Decompress))
                    {
                        gzip.CopyTo(outputStream);
                    }
                }
            }
        }
        private static string GDriveCreateFolder(DriveService service, string foldername)
        {
            var file = new File();
            file.Name = foldername;
            file.MimeType = "application /vnd.google.app.folder";

            var request = service.Files.Create(file);
            request.Fields = "id";

            var result = request.Execute();
            return result.Id;
        }
        private static UserCredential GetUserCredentials()
        {
            using (var stream = new FileStream("client_secret_.json", FileMode.Open, FileAccess.Read))
            {
                var creadPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                creadPath = Path.Combine(creadPath, "driveApiCredantials", "drive-credentials.json");
                   

                return GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    scope,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(creadPath, true)).Result;
            }
        }
        private static DriveService GetDriveService(UserCredential credential)
        {
            return new DriveService(
                new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = applicationName
                });
        }

        public static void DelayAndQuit(int code =0)
        {
            Thread.Sleep(15000);
            Environment.Exit(code);
        } 
    }
}
