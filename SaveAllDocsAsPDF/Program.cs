using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DriveQuickstart
{


   
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/drive-dotnet-quickstart.json
        static string[] Scopes = { DriveService.Scope.Drive };
        static string ApplicationName = "Drive API .NET Quickstart";
        static bool RanOnce = false;
        static string SubfolderName = "ResavedPDFs";

        static string GetOrCreateFolder (DriveService service, string SuperFolder)
        {
            Google.Apis.Drive.v3.Data.File SubFolder = GetFolderInSubfolder(service, SuperFolder);


            if (SubFolder?.Id != null)
            {
                return SubFolder.Id;
            }
            else
            {
                var FolderMetadata = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = SubfolderName,
                    MimeType = "application/vnd.google-apps.folder"
                };
                if (string.IsNullOrEmpty(SuperFolder))
                {
                    //do nothing
                }
                else
                {
                    FolderMetadata.Parents = new List<string>();
                    FolderMetadata.Parents.Add(SuperFolder);
                }
                var request = service.Files.Create(FolderMetadata);
                request.Fields = "id";

                var folder = request.Execute();
                Console.WriteLine("Folder ID: " + folder.Id);

                return folder.Id;
            }
        }

        static Google.Apis.Drive.v3.Data.File GetFolderInSubfolder(DriveService driveService , string parent)
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
                            if (file.Name.Equals(SubfolderName))
                            {
                                return file;
                            }

                        }
                        else
                        {
                            if (file.Name.Equals(SubfolderName) && file.Parents.Contains(parent))
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


        static void Main(string[] args)
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
                CheckAndUpdate();
                Thread.Sleep(minutes * 1000 * 60);
            } while (WantToKeepRunning | (RanOnce == false));
        }

        static bool CheckAndUpdate()
        {
            UserCredential credential;
            Console.WriteLine("Started running at: " + DateTime.Now.TimeOfDay);
            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart2.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            string LastModifiedUnixDate = "LastRan.txt";
            DateTime DateTimeProgramLastRan = new DateTime(0);

            string LastModifiedToken = "token.txt";
            string TokenData = "";
            try
            {
                using (var ProgramInfo = new StreamReader(LastModifiedToken, Encoding.UTF8))
                {
                    TokenData = ProgramInfo.ReadLine();
                }
            }
            catch (Exception ex)
            { }

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
            

            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.

            List<Google.Apis.Drive.v3.Data.File> AllFiles = new List<Google.Apis.Drive.v3.Data.File>();
            
            if (string.IsNullOrEmpty(TokenData) == false)
            {
                ChangesResource.ListRequest ModifiedOnlyList = service.Changes.List(TokenData);
                string NextPageToken = "";
                do
                {

                    var execute = ModifiedOnlyList.Execute();
                    //  ModifiedOnlyList.Fields = "nextPageToken, files(id, mimeType, name, parents, modifiedTime  )";
                    ChangesResource item;

                    execute = ModifiedOnlyList.Execute();
                    NextPageToken = execute.NextPageToken;

                    foreach (var entry in execute.Changes)
                    {
                        if (entry.File != null)
                        {
                            AllFiles.Add(entry.File);
                        }

                    }
                    ModifiedOnlyList = service.Changes.List(execute.NextPageToken);

                } while (string.IsNullOrEmpty(NextPageToken) == false);
                
            }

            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 1000;
            listRequest.Fields = "nextPageToken, files(id, mimeType, name, parents, modifiedTime  )";
            listRequest.OrderBy = "modifiedTime desc";

            // List files.
            FileList ListAll;
            List<Google.Apis.Drive.v3.Data.File> AllExistingFiles = new List<Google.Apis.Drive.v3.Data.File>();
            do
            {
                ListAll = listRequest.Execute();
                IList<Google.Apis.Drive.v3.Data.File> files = ListAll.Files;
                
                foreach (var file in files)
                {
                    if (file.ModifiedTime.Value > DateTimeProgramLastRan)
                    {
                        if (file != null) {
                            AllFiles.Add(file);
                        }
                    }
                    AllExistingFiles.Add(file);
                }
                listRequest.PageToken = ListAll.NextPageToken;

            } while (string.IsNullOrEmpty(ListAll.NextPageToken) == false);

            
            
            

            foreach (var file in AllFiles)
            {
                string GoogleDoc = "application/vnd.google-apps.document";

                if (file.MimeType.Equals(GoogleDoc) && (file?.Parents != null))
                {
                    Console.WriteLine("Change found in: {0}",file.Name);

                    string appendtofile = "_Resaved.pdf";

                    Google.Apis.Drive.v3.Data.File fileMetadata = new Google.Apis.Drive.v3.Data.File()
                    {
                        Name = file.Name + appendtofile,
                        Parents = file.Parents
                    };
                    if (file?.Parents?.Count == null)
                    {
                        fileMetadata.Parents = new List<string>();
                        fileMetadata.Parents.Add(GetOrCreateFolder(service, null));
                      
                        
                    }
                    else
                    {
                        fileMetadata.Parents = new List<string>();
                        fileMetadata.Parents.Add(GetOrCreateFolder(service, file.Parents[0]));
                    }
                    

                    // convert string to stream

                    bool UpdateFile = true;

                    
                    
                    foreach (var ExistingFile in AllExistingFiles)
                    {
                        string samplename = file.Name + appendtofile;

                        if (samplename.Equals(ExistingFile.Name))
                        {
                            bool CanDoComparison = true;
                            CanDoComparison = ((file?.Parents != null) );
                            if (CanDoComparison)
                            {
                                string PDFID = GetOrCreateFolder(service, file.Parents[0]);

                                if (ExistingFile.Parents.Contains(PDFID))
                                {
                                    try
                                    {
                                        service.Files.Delete(ExistingFile.Id).Execute();
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("There was an error deleting: {0} with id: {1}", ExistingFile.Name, ExistingFile.Id);
                                    }
                                    finally
                                    {
                                        UpdateFile = true;
                                    }
                                }
                            }
                        } 
                    }
                    
                    //Check this isn't a file that is currently open and being edited, wait till edits are made before resaving
                    foreach (var fileRHS in AllFiles)
                    {
                        if (file.Id.Equals(fileRHS.Id) && file?.Parents == null)
                        {
                            Console.WriteLine(file.Name + "Is currently being modified, we will wait until changes have been completed");
                            UpdateFile = false;
                            break;
                        }
                    }
                    if (UpdateFile)
                    {
                        //Create New File
                        var stream = new System.IO.MemoryStream();
                        service.Files.Export(file.Id, "application/pdf").Download(stream);
                        using (stream)
                        {
                            FilesResource.CreateMediaUpload request;
                            request = service.Files.Create(fileMetadata, stream, "application/pdf");

                            request.Fields = "id";
                            request.Upload();
                            var item = request.ResponseBody;
                        }
                    }

                    using (System.IO.StreamWriter NewFileWrite = new System.IO.StreamWriter(LastModifiedToken, false))
                    {
                        var response = service.Changes.GetStartPageToken().Execute();
                        NewFileWrite.WriteLine(response.StartPageTokenValue);
                    }

                    using (System.IO.StreamWriter NewFileWrite = new System.IO.StreamWriter(LastModifiedUnixDate, false))
                    {
                        NewFileWrite.WriteLine(DateTime.Now.Ticks);
                    }
                   
                }
            }
            Console.WriteLine("Finished updating");
            RanOnce = true;
            return true;
        }
    }
}
