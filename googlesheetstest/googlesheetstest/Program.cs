using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Spreadsheets;
using Newtonsoft.Json;
using Data = Google.Apis.Sheets.v4.Data;


namespace googlesheetstest
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        //Get And Set Spread sheet data
        static string[] Scopes = { SheetsService.Scope.Drive};
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        static void Main(string[] args)
        {
            CreateNewSheet();
        }
        static UserCredential GetCredential()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("C:/Users/chartsuff/Downloads/client_id.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }
        static void CreateNewSheet()
        {
            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = ApplicationName,
            });

            // TODO: Assign values to desired properties of `requestBody`:
            Data.Spreadsheet requestBody = new Data.Spreadsheet();
            //IDs
            requestBody.SpreadsheetId = "thebakerid";
            Sheet sheet = new Sheet();
            Sheet sheet2 = new Sheet();
            sheet.Properties = new SheetProperties();               
            sheet2.Properties = new SheetProperties();
            sheet.Properties.Title = "test";
            sheet.Properties.SheetId = 1;
            sheet2.Properties.SheetId = 2;
            sheet.Properties.Title = "test2";
            requestBody.Sheets = new List<Sheet>();
            requestBody.Sheets.Add(sheet);
            requestBody.Sheets.Add(sheet2);
            //



            SpreadsheetsResource.CreateRequest request = sheetsService.Spreadsheets.Create(requestBody);

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Data.Spreadsheet response = request.Execute();
            // Data.Spreadsheet response = await request.ExecuteAsync();

            // TODO: Change code below to process the `response` object:
            Console.WriteLine(JsonConvert.SerializeObject(response));
        }
        static void GetSpreadData()
        {          
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = ApplicationName,
            });

            ValueRange writestuff = new ValueRange().Values = ;

            // Define request parameters.
            String spreadsheetId = "1iZ5lzJV1qYlhoY-n1rCP8cPmpGQOoHcjBKSmviq47-8";
            String range = "Sheet1!A2:D3";
            SpreadsheetsResource.ValuesResource.UpdateRequest request =
                    service.Spreadsheets.Values.Update(writestuff ,spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            request.Execute();
            /*ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    foreach (var col in row)
                    {
                        Console.WriteLine(col);
                    }
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            */
            Console.Read();
        }
    }
}
