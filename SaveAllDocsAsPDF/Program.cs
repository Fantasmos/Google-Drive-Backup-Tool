using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace DriveQuickstart
{
    internal class Program
    {
        private static string ApplicationName = "Drive API .NET Quickstart";

        private static string LastModifiedToken = "token.txt";
        private static string LastModifiedUnixDate = "LastRan.txt";
        private static bool RanOnce = false;

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/drive-dotnet-quickstart.json
        private static string[] Scopes = {
   DriveService.Scope.Drive
  };

        public static List<Google.Apis.Drive.v3.Data.File> Get_All_Files(DriveService service, string tokendata)
        {
            // Define parameters of request.
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 1000;
            listRequest.Fields = "nextPageToken, files(id, mimeType, name, parents, modifiedTime  )";
            listRequest.OrderBy = "modifiedTime desc";

            List<Google.Apis.Drive.v3.Data.File> AllExistingFiles = new List<Google.Apis.Drive.v3.Data.File>();
            FileList ListAll;
            do
            {
                ListAll = listRequest.Execute();
                IList<Google.Apis.Drive.v3.Data.File> files = ListAll.Files;

                foreach (var file in files)
                {
                    AllExistingFiles.Add(file);
                }
                listRequest.PageToken = ListAll.NextPageToken;
            } while (string.IsNullOrEmpty(ListAll.NextPageToken) == false);
            return AllExistingFiles;
        }

        private static void CheckAndUpdate(DriveService service, List<Google.Apis.Drive.v3.Data.File> All_Files, DateTime DateTimeProgramLastRan)
        {
            List<Conversions> AllConversions = GetAllConversions();

            foreach (var file in All_Files)
            {
                foreach (Conversions conversion in AllConversions)
                {
                    //If File is Valid and needs to be updated
                    if ((file.ModifiedTime.Value > DateTimeProgramLastRan) && file.MimeType.Equals(conversion.NativeFormat) && (file?.Parents != null))
                    {

                        Console.WriteLine("Change found in: {0}", file.Name);

                        foreach (format OutputFormat in conversion.Outputs)
                        {

                            string Resaves_FolderID = file.Parents[0];

                            var fileMetadata = new Google.Apis.Drive.v3.Data.File();
                            fileMetadata.Name = file.Name + OutputFormat.name;
                            fileMetadata.Parents = new List<string>() { Resaves_FolderID };

                            var Existing_File = Get_File_By_Name_And_SuperFolder(Resaves_FolderID, fileMetadata.Name, All_Files);

                            Delete_Existing_File(service, Existing_File);

                            Convert_and_Resave_File(service, file, fileMetadata, OutputFormat.MimeHeader);
                        }
                    }
                }
            }
        }

        private static string Create_Folder(DriveService service, string Location, string FolderName)
        {
            var FolderMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = FolderName,
                MimeType = "application/vnd.google-apps.folder"
            };

            if (string.IsNullOrEmpty(FolderName))
            {
                //do nothing
            }
            else
            {
                FolderMetadata.Parents = new List<string>();
                FolderMetadata.Parents.Add(Location);
            }
            var request = service.Files.Create(FolderMetadata);
            request.Fields = "id";

            var folder = request.Execute();
            Console.WriteLine("Folder ID: " + folder.Id);

            return folder.Id;
        }

        private static void Delete_Existing_File(DriveService service, Google.Apis.Drive.v3.Data.File ExistingFile)
        {
            if (ExistingFile != null)
            {
                try
                {
                    service.Files.Delete(ExistingFile.Id).Execute();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("There was an error deleting: {0} with id: {1}", ExistingFile.Name, ExistingFile.Id);
                }
            }
        }

        private static Conversions GDocs()
        {
            format pdf = new format();
            pdf.name = "pdf";
            pdf.MimeHeader = "application/pdf";

            format OpenDocument = new format();
            OpenDocument.name = "odt";
            OpenDocument.MimeHeader = "application/vnd.oasis.opendocument.text";

            

            format[] AllFormats =
                {
                    pdf,
                    OpenDocument
                };

            Conversions GoogleDoc_To_ODT_and_PDF;

            GoogleDoc_To_ODT_and_PDF.NativeFormat = "application/vnd.google-apps.document";
            GoogleDoc_To_ODT_and_PDF.Outputs = AllFormats;


            format docx = new format();
            docx.name = ".docx";
            docx.MimeHeader = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            format[] AllFormats2 =
                {
                    docx
                };

            Conversions GDOC_TO_DOCX;
            GDOC_TO_DOCX.NativeFormat = "application/vnd.google-apps.document";
            GDOC_TO_DOCX.Outputs = AllFormats2;

            return GDOC_TO_DOCX;
        }

        private static List<Google.Apis.Drive.v3.Data.File> Get_All_Modified_Files(string TokenData, DriveService service)
        {
            List<Google.Apis.Drive.v3.Data.File> All_Modified_Files = new List<Google.Apis.Drive.v3.Data.File>();
            if (string.IsNullOrEmpty(TokenData) == false)
            {
                ChangesResource.ListRequest Modified_Only_List = service.Changes.List(TokenData);
                string NextPageToken = "";
                do
                {
                    var execute = Modified_Only_List.Execute();
                    ChangesResource item;

                    execute = Modified_Only_List.Execute();
                    NextPageToken = execute.NextPageToken;

                    foreach (var entry in execute.Changes)
                    {
                        if (entry.File != null)
                        {
                            All_Modified_Files.Add(entry.File);
                        }
                    }
                    Modified_Only_List = service.Changes.List(execute.NextPageToken);
                } while (string.IsNullOrEmpty(NextPageToken) == false);
            }
            return All_Modified_Files;
        }

        private static Google.Apis.Drive.v3.Data.File Get_File_By_Name_And_SuperFolder(string ParentID, string filename, List<Google.Apis.Drive.v3.Data.File> ExistingFiles, bool MustBeFolder = false)
        {
            foreach (var ExistingFile in ExistingFiles)
            {
                if (ExistingFile.Name.Equals(filename) && ExistingFile.Parents.Contains(ParentID))
                {
                    //Future Implementation
                    if (MustBeFolder && ExistingFile.MimeType.Equals("application/vnd.google-apps.folder"))
                    {
                        return ExistingFile;
                    }
                    else if (MustBeFolder == false)
                    {
                        return ExistingFile;
                    }
                }
            }
            return null;
        }

        private static Google.Apis.Drive.v3.Data.File Get_Folder(DriveService driveService, string parent, string FolderName)
        {
            try
            {
                string pageToken = null;
                do
                {
                    var request = driveService.Files.List();
                    request.Q = "mimeType='application/vnd.google-apps.folder'";
                    request.Spaces = "drive";
                    request.Fields = "nextPageToken, files(id, name, parents)";
                    request.PageToken = pageToken;
                    var result = request.Execute();
                    foreach (var file in result.Files)
                    {
                        if (string.IsNullOrEmpty(parent) | file?.Parents == null)
                        {
                            if (file.Name.Equals(FolderName))
                            {
                                return file;
                            }
                        }
                        else
                        {
                            if (file.Name.Equals(FolderName) && file.Parents.Contains(parent))
                            {
                                return file;
                            }
                        }
                    }
                    pageToken = result.NextPageToken;
                } while (pageToken != null);

                return null;
            }
            catch
            {
                Console.WriteLine("Threw error");
                return null;
            }
        }

        private static string Get_Or_Create_Folder(DriveService service, string SuperFolder, string Folder_Name)
        {
            Google.Apis.Drive.v3.Data.File Folder = Get_Folder(service, SuperFolder, Folder_Name);
            
            if (Folder?.Id != null)
            {
                return Folder.Id;
            }
            else
            {
                return Create_Folder(service, SuperFolder, Folder_Name);
            }
        }

        private static format[] GetAllFormats()
        {
            format pdf = new format();
            pdf.name = "pdf";
            pdf.MimeHeader = "application/pdf";

            format OpenDocument = new format();
            OpenDocument.name = "odt";
            OpenDocument.MimeHeader = "application/vnd.oasis.opendocument.text";

            format[] AllFormats = {
    pdf,
    OpenDocument
   };
            return AllFormats;
        }

        private static UserCredential GetCredential()
        {
            Console.WriteLine("Started running at: " + DateTime.Now.TimeOfDay);
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart2.json");

                UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                 GoogleClientSecrets.Load(stream).Secrets,
                 Scopes,
                 "user",
                 CancellationToken.None,
                 new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
                return credential;
            }
        }

        private static DateTime GetLastTimeRan()
        {
            DateTime DateTimeProgramLastRan = new DateTime(0);
            try
            {
                using (var ProgramInfo = new StreamReader(LastModifiedUnixDate, Encoding.UTF8))
                {
                    string FileContents = ProgramInfo.ReadToEnd();
                    long Ticks = long.Parse(FileContents);
                    DateTimeProgramLastRan = new DateTime(Ticks);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("When this program attempted to read/load {0} an error occured! The program will not filter files that have not been changed.", LastModifiedUnixDate);
            }
            return DateTimeProgramLastRan;
        }

        private static string GetTokenData()
        {
            string TokenData = "";

            try
            {
                using (var ProgramInfo = new StreamReader(LastModifiedToken, Encoding.UTF8))
                {
                    TokenData = ProgramInfo.ReadLine();
                }
            }
            catch (Exception ex) { }
            return TokenData;
        }

        public static bool Run()
        {
            UserCredential credential = GetCredential();
            string tokendata = GetTokenData();

            DateTime DateTimeProgramLastRan = GetLastTimeRan();
            DriveService service = CreateService(credential, ApplicationName);

            var Current_Drive_Files = Get_All_Files(service, tokendata);

            CheckAndUpdate(service, Current_Drive_Files, DateTimeProgramLastRan);

            Save_StartPage_Token(service);
            Save_Current_Time();
            Console.WriteLine("Finished updating");
            RanOnce = true;
            return true;
        }

        private static void Save_Current_Time()
        {
            using (System.IO.StreamWriter NewFileWrite = new System.IO.StreamWriter(LastModifiedUnixDate, false))
            {
                NewFileWrite.WriteLine(DateTime.Now.Ticks);
            }
        }

        private static void Save_StartPage_Token(DriveService service)
        {
            using (System.IO.StreamWriter NewFileWrite = new System.IO.StreamWriter(LastModifiedToken, false))
            {
                var response = service.Changes.GetStartPageToken().Execute();
                NewFileWrite.WriteLine(response.StartPageTokenValue);
            }
        }

        private static DriveService CreateService(UserCredential credential, string ApplicationName)
        {
            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        private static List<Conversions> GetAllConversions()
        {
            Conversions GoogleDoc_To_ODT_and_PDF = GDocs();

            List<Conversions> AllConversions = new List<Conversions>();
            AllConversions.Add(GoogleDoc_To_ODT_and_PDF);
            return AllConversions;
        }

        public struct Conversions
        {
            public string NativeFormat;
            public format[] Outputs;
        }

        public struct format
        {
            public string MimeHeader;
            public string name;
        }

        private static void Convert_and_Resave_File(DriveService service, Google.Apis.Drive.v3.Data.File file, Google.Apis.Drive.v3.Data.File fileMetadata, string metadatatype)
        { //Create New File
            var stream = new System.IO.MemoryStream();
            service.Files.Export(file.Id, metadatatype).Download(stream);
            using (stream)
            {
                FilesResource.CreateMediaUpload request;
                request = service.Files.Create(fileMetadata, stream, metadatatype);

                request.Fields = "id";
                request.Upload();
                var item = request.ResponseBody;
            }
        }

        private static void Main(string[] args)
        {
            bool WantToKeepRunning = true;
            if (args.Length > 0)
            {
                WantToKeepRunning = bool.Parse(args[0]);
            }
            int minutes = 5;

            if (args.Length > 1)
            {
                minutes = int.Parse(args[1]);
            }

            Console.WriteLine("Timer set to: " + minutes);

            do
            {

                Run();
                Thread.Sleep(minutes * 1000 * 60);
            } while (WantToKeepRunning | (RanOnce == false));
        }
    }
}
 
