using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using System;
using System.Configuration;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using Data = Google.Apis.Sheets.v4.Data;

namespace NewForma_Log_Command
{
    public class GoogleSheetsClass
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets Logger";

        public static void WritetoGoogleSheets(string tabname, IList<IList<object>> DataToWrite)
        {
            UserCredential credential;

            using (var stream = 
                //new FileStream(@"C:\Users\acarpentier\source\repos\Optimus NewForma Log\Procore Log Command\bin\Release\net5.0\publish\credentials.json", FileMode.Open, FileAccess.ReadWrite))
                new FileStream(ConfigurationManager.AppSettings.Get("GoogleSheetsCredentials"), FileMode.Open, FileAccess.ReadWrite))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = ConfigurationManager.AppSettings.Get("GoogleSheetsToken");
                //string credPath = @"C:\Users\acarpentier\source\repos\Optimus NewForma Log\Procore Log Command\bin\Release\net5.0\publish\token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            string spreadsheetId = ConfigurationManager.AppSettings.Get("spreadsheetId");

            // empty table
            IList<IList<object>> updateData = new List<IList<object>>();
            Data.ValueRange requestBody = new Data.ValueRange();
            requestBody.Range = $"{tabname}!A4";
            requestBody.Values = updateData;

            SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueInputOption = (SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum)2;

            // append the empty table to the spreadsheet
            SpreadsheetsResource.ValuesResource.AppendRequest request = service.Spreadsheets.Values.Append(requestBody, spreadsheetId, $"{tabname}!A4");
            request.ValueInputOption = valueInputOption;
            Data.AppendValuesResponse response = request.Execute();

            //the table is appended to the end of the spreadsheet, allowing us to find the first empty row
            string responsestring = JsonConvert.SerializeObject(response);
            Dictionary<string, object> responsedict = JsonConvert.DeserializeObject<Dictionary<string, object>>(responsestring);
            string jsonrange = JsonConvert.SerializeObject(responsedict["tableRange"]);

            int emptyRowIndex = Convert.ToInt32(Regex.Match(jsonrange, @"![A-Z]+\d+:[A-Z]+(\d+)").Groups[1].Value) + 1;
            string emptyRowRange = $"{tabname}!A{emptyRowIndex}";

            //set the filled list to be added in the first empty row
            requestBody = new Data.ValueRange();
            requestBody.Range = emptyRowRange;
            //requestBody.Range = $"{tabname}!A4";
            requestBody.Values = DataToWrite;

            SpreadsheetsResource.ValuesResource.UpdateRequest request2 = service.Spreadsheets.Values.Update(requestBody, spreadsheetId, emptyRowRange);
            SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum valueInputOption2 = (SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum)2;
            request2.ValueInputOption = valueInputOption2;

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Data.UpdateValuesResponse response2 = request2.Execute();

        }
    }
}