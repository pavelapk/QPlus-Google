using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;

namespace test_google_api;

internal static class Program {
    private static readonly string[] Scopes = {SheetsService.Scope.Spreadsheets};

    private const string CredentialsFilename = "google_credentials.json";

    private static ServiceAccountCredential GetCredential() {
        var credentialParameters = GetCredentialParameters();
        if (credentialParameters == null) {
            throw new FileNotFoundException(CredentialsFilename);
        }

        return new ServiceAccountCredential(
            new ServiceAccountCredential.Initializer(credentialParameters.ClientEmail) {
                Scopes = Scopes,
                // User = User
            }.FromPrivateKey(credentialParameters.PrivateKey));
    }

    private static JsonCredentialParameters? GetCredentialParameters() {
        return JsonConvert.DeserializeObject<JsonCredentialParameters>(File.ReadAllText(CredentialsFilename));
    }

    private static void Main(string[] args) {
        var credential = GetCredential();

        // Create Google Sheets API service.
        var service = new SheetsService(new BaseClientService.Initializer {
            HttpClientInitializer = credential,
            // ApplicationName = ApplicationName,
        });

        // Define request parameters.
        var spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
        var range = "Class Data!A2:E";
        var request = service.Spreadsheets.Values.Get(spreadsheetId, range);

        // Prints the names and majors of students in a sample spreadsheet:
        // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
        var response = request.Execute();
        IList<IList<object>> values = response.Values;
        if (values != null && values.Count > 0) {
            Console.WriteLine("Name, Major");
            foreach (var row in values) {
                // Print columns A and E, which correspond to indices 0 and 4.
                Console.WriteLine("{0}, {1}", row[0], row[4]);
            }
        } else {
            Console.WriteLine("No data found.");
        }

        Console.Read();
    }
}