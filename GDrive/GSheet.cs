using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Google.Apis.Util;
using RigLauncher;

namespace GDrive
{
    public static class GSheet
    {
        public static int row;
        private static readonly string[] ScopesSheets = { SheetsService.Scope.Spreadsheets };
        public static readonly string ClientSecret = "client_secret_.json";
        public static readonly string SpreadsheetId = GetSpreadsheetId();
        private static int biosId ;
        private static UserCredential credential;
        private static SheetsService service;
        private static SpreadsheetsResource.ValuesResource.BatchGetRequest request;
        private static BatchGetValuesResponse response;
        private static readonly string  myName = Environment.MachineName.ToLower();
        public static readonly string version = "version";
        private static int newVersion;
        private static int rowVersion;
        static GSheet()
        {
            credential = GetSheetCredentials();
            service = GetService(credential);
            biosId = GetBiosId("bios");
        }

        private static int GetBiosId(string spreadsheetId)
        {
            int result = 0;
            var meta = service.Spreadsheets.Get(SpreadsheetId).Execute();
            var bioSheet = meta.Sheets.FirstOrDefault(i => i.Properties.Title == spreadsheetId);
            if (bioSheet != null && !int.TryParse(bioSheet.Properties.SheetId.ToString(), out result))
            {
                Console.WriteLine("Cannot find biosID");
            }
            return result;
        }

        public static RigVersion ReadGSheet()
        {
            var Firstline = GetRange("bios!1:1");
            string zipId = Firstline[0][3].ToString();
            int newVersion;
            if (!int.TryParse(Firstline[0][2].ToString(), out newVersion))
            {
                RigEx.WriteLineColors("Google Sheet => Cannot read new version on C1 ", ConsoleColor.Red);
                return null;
            }
            var range = Firstline[0][0].ToString();
            var res = GetRange(range);

            for (int i = 0; i < res.Count; i++)
            {
                var line = res[i];
                if (line.Count == 0 || string.IsNullOrEmpty(line[0].ToString()))
                    continue;

                string attribute = line[0].ToString().ToLower();
                if (attribute == version)
                {
                    if (line[1].ToString().ToLower() != myName)
                        continue;
                    int curVersion;
                    if (line != null && line.Count >= 3
                        && int.TryParse(line[2].ToString(), out curVersion))
                    {
                        Console.WriteLine("row: "+i);
                        rowVersion = i;
                        return new RigVersion(curVersion,newVersion, zipId);
                    }
                }
                
            }
            return null;

        }
        private static UserCredential GetSheetCredentials()
        {
            var path = Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetEntryAssembly().Location, @"..\"));
            try
            {
                using (var stream = new FileStream(Path.Combine(path, ClientSecret), FileMode.Open,
                    FileAccess.Read))
                {

                    var creadPath = Path.Combine(path, "sheetsCreds.json");

                    return GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                        ScopesSheets, "user", CancellationToken.None, new FileDataStore(creadPath, true)).Result;
                }
            }
            catch (Exception e)
            {
                RigEx.WriteLineColors($"Cannot read Google Sheet secret key Error:{e.Message}".AddTimeStamp(), ConsoleColor.Red);
                Launcher.DelayAndQuit();
                return null;
            }
        }

        private static SheetsService GetService(UserCredential credential)
        {
            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "ForRig"
            });
        }

        private static IList<IList<object>> GetRange(params string[] range)
        {
            request = service.Spreadsheets.Values.BatchGet(SpreadsheetId);
            request.Ranges = new Repeatable<string>(range);
            IList<IList<object>> result;
            try
            {
                response = request.Execute();
                result = response.ValueRanges.SelectMany(x => x.Values).ToList();
                return result;
            }
            catch (Exception)
            {
                Console.WriteLine($"Google Sheet read error from {range[0]}");
                Launcher.DelayAndQuit();
                return null;
            }
        }

        public static  void SendData(string value)
        {
           List<Request> request = new List<Request>
            {
                GetRequest(value)
            };
            SendRequests(request);
        }

        private static  Request GetRequest(string value)
        {
            List<CellData> values = new List<CellData>();
            Request result;

            values.Add(new CellData
                {
                    
                    UserEnteredValue = new ExtendedValue
                    {
                        StringValue = value
                    }

                });

                result = new Request
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Start = new GridCoordinate
                        {
                            
                            SheetId = biosId,
                            RowIndex = rowVersion,
                            ColumnIndex = 2
                        },
                        Rows = new List<RowData> { new RowData { Values = values } },
                        Fields = "UserEnteredFormat(BackgroundColor),userEnteredValue"
                    }
                };
            
            return result;
        }

        private static void SendRequests(List<Request> request)
        {
            var busr = new BatchUpdateSpreadsheetRequest { Requests = request };
            try
            {
                var t = service.Spreadsheets.BatchUpdate(busr, SpreadsheetId);
                //t.Execute();
                t.ExecuteAsStream();
            }
            catch (Exception e)
            {
                RigEx.WriteLineColors($"Exeption to send request{e.Message}".AddTimeStamp(), ConsoleColor.Red);

            }

        }
        private static string GetSpreadsheetId()
        {
            string path = Path.GetFullPath(Path.Combine(Path.Combine(System.Reflection.Assembly.GetEntryAssembly().Location, @"..\"), "sheetId.txt"));
            //string path = @"c:\temp\MyTest.txt";
            if (!File.Exists(path))
            {
                RigEx.WriteLineColors($"Cannot faund file {path} ", ConsoleColor.Red);
                return String.Empty;
            }
            var t = File.ReadAllText(path);
            return t;
        }
    }
}
