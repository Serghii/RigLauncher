using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using File = Google.Apis.Drive.v3.Data.File;

namespace GDrive
{
    public static class GDrive
    {
        public static Action DownloadCompletedAction = () => { Console.WriteLine("Download Completed"); };

        private static string[] scope = { DriveService.Scope.Drive };
        private static string applicationName = "GDrive";
        private static string FolderPathId = "1_5-j9pPV0auB3cvSVqSTd1j_sWtIRXEE";
        private static string contentType = "application/zip";
        private static DriveService service;
       static GDrive()
        {
            UserCredential credential = GetUserCredentials();
            service = GetDriveService(credential);
        }
        public static string UploadFileToGDrive( string fileName,string filePath)
        {
            var fileMetadata = new File();
            fileMetadata.Name = fileName;
            fileMetadata.Parents = new List<string> { FolderPathId };

            FilesResource.CreateMediaUpload request;

            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                request = service.Files.Create(fileMetadata, stream, contentType);
                request.Upload();
            }

            var file = request.ResponseBody;
            
            return file.Id;
        }

        public static void DownloadFileFromGH(string fileId, string filePath)
        {
            var lastReleas = "https://api.github.com/<Serghii>/<Rig>/releases/latest";

        }
        public static void DownloadFileFromGDrive( string fileId, string filePath)
        {
            Console.WriteLine($"Google Drive Begin DownLoad to {filePath}");
            var request = service.Files.Get(fileId);
            using (var memoryStream = new MemoryStream())
            {
                request.MediaDownloader.ProgressChanged += (IDownloadProgress progress) =>
                {
                    switch (progress.Status)
                    {
                        case DownloadStatus.NotStarted:
                            Console.WriteLine("Download Not Started");
                            break;
                        case DownloadStatus.Downloading:
                            Console.WriteLine($"Download: {progress.BytesDownloaded}");
                            break;
                        case DownloadStatus.Completed:
                            Console.WriteLine("Download Completed");
                            break;
                        case DownloadStatus.Failed:
                            Console.WriteLine("Download Failed");
                            break;
                        default:
                            Console.WriteLine("Download default");
                            break;
                    }
                };
                request.Download(memoryStream);

                try
                {
                    using (var fileStrim = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        fileStrim.Write(memoryStream.GetBuffer(), 0, memoryStream.GetBuffer().Length);
                        fileStrim.Close();
                    }
                }
                catch (Exception e)
                {
                    RigEx.WriteLineColors($"Can not save zip to {filePath} Error: {e.Message}",ConsoleColor.DarkRed);
                }
                 DownloadCompletedAction();
            }

        }
        public static string GDriveCreateFolder(string foldername)
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
            var path = Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetEntryAssembly().Location, @"..\"));

           using (var stream = new FileStream(Path.Combine(path, GSheet.ClientSecret), FileMode.Open, FileAccess.Read))
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
    }
}
