using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Microsoft.Extensions.Configuration;

namespace GoogleApiLib
{
    public class GoogleSheetService
    {
        private string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        private IConfiguration _config;

        public GoogleSheetService(IConfiguration config)
        {
            _config = config;
        }

        public IList<IList<object>> GetGoogleSheetContent(
            string credentialPath,
            string sheetId, 
            string tabName, 
            string range = "")
        {
            range = $"{tabName}!{range}";

            using (var stream =
                new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
                };

                var service = new SheetsService(serviceInitializer);
                SpreadsheetsResource.ValuesResource.GetRequest request =
                  service.Spreadsheets.Values.Get(sheetId, range);

                ValueRange response = request.Execute();
                return response.Values;
            }
        }
    }
}
