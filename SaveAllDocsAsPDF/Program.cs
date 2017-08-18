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

namespace DriveQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/drive-dotnet-quickstart.json
        static string[] Scopes = { DriveService.Scope.Drive };
        static string ApplicationName = "Drive API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

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
                }
                listRequest.PageToken = ListAll.NextPageToken;

            } while (string.IsNullOrEmpty(ListAll.NextPageToken) == false);

            
            Console.WriteLine("Files:");
            

            foreach (var file in AllFiles)
            {
                string GoogleDoc = "application/vnd.google-apps.document";

                if (file.MimeType.Equals(GoogleDoc))
                {
                    Console.WriteLine("{0}", file.MimeType);
                    string appendtofile = "_Resaved.pdf";
                    Google.Apis.Drive.v3.Data.File fileMetadata = new Google.Apis.Drive.v3.Data.File()
                    {
                        Name = file.Name + appendtofile,
                        Parents = file.Parents
                    };

                    // convert string to stream

                    bool UpdateFile = false;

                    
                    
                    foreach (var ExistingFile in AllFiles)
                    {
                        string samplename = file.Name + appendtofile;
                        if (samplename.Equals(ExistingFile.Name))
                        {
                            bool CanDoComparison = true;

                            CanDoComparison = ((file?.Parents != null) & (ExistingFile?.Parents != null));

                            //If they're both null, or both are not null, we keep executing 
                            if (CanDoComparison) {
                                if (Enumerable.SequenceEqual(file.Parents, ExistingFile.Parents))
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

                        Console.WriteLine("Start token: " + response.StartPageTokenValue);

                        NewFileWrite.WriteLine(response.StartPageTokenValue);
                    }
                    Console.WriteLine(file.Name);
                }

               

            }
        }
    }
}
