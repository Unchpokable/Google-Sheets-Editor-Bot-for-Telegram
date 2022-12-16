using GSheetsEditor.Commands.Assembly;
using Google.Apis.Sheets;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using System.Text.RegularExpressions;
using System.Text;
using GSheetsEditor.Extensions;
using Google.Apis.Sheets.v4.Data;

namespace GSheetsEditor.Commands.Modules
{
    [CommandModule]
    internal class GoogleAPICommands
    {
        private static readonly string[] _scopes = { SheetsService.Scope.Spreadsheets };
        private static readonly string _applicationName = "app";
        private static string _defaultSpreadsheetID = "1wG0yquhycja9AWrf1xjvCbkDW5UA2BTPtQqeyKhzST0";
        private static readonly string _baseSheet = "Sheet1";
        private static SheetsService _service;

        static GoogleAPICommands()
        {
            GoogleCredential credentials;

            using (var file = new FileStream("GoogleApiCredentials.json", FileMode.Open, FileAccess.Read))
            {
                credentials = GoogleCredential.FromStream(file)
                                .CreateScoped(_scopes);
            }

            _service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credentials,
                ApplicationName = _applicationName,
            });
        }

        [Command("/read")]
        public static CommandExecutionResult ReadData(object parameter)
        {
            if (!ValidateCell(parameter, out CommandExecutionResult errorMessage))
                return errorMessage;

            var cell = parameter.ToString();

            try
            {
                var value = _service.Spreadsheets.Values.Get(_defaultSpreadsheetID, $"{_baseSheet}!{cell}").Execute().Values;
                if (value == null)
                    return new CommandExecutionResult("Empty Cell");
                return new CommandExecutionResult(TableFetchedRangeToString(value, cell[0], int.Parse(cell.From(1).To(':'))));
            }
            catch (Exception ex)
            {
                return new CommandExecutionResult($"ReadValue method exception thrown: {ex.Message}");
            }
        }

        [Command("/write")]
        public static CommandExecutionResult WriteData(object parameter)
        {
            // CommandArgs awaited structure: string`[range, ..args <- at least one]
            if (parameter is not IList<string> commandArgs)
            {
                return new CommandExecutionResult("Invalid argument. Usage: /write {range} {list_of_values} for multiple values or /write {cell} {single_value} for write single value in specified cell");
            }

            if (commandArgs.Count == 0 || commandArgs.Count == 1)
            {
                return new CommandExecutionResult("Invalid argument. Usage: /write {range} {list_of_values} for multiple values or /write {cell} {single_value} for write single value in specified cell");
            }

            if (!ValidateCell(commandArgs[0], out CommandExecutionResult errorMessage)) 
                return errorMessage;

            var valuesToWrite = commandArgs.Skip(1).Select(v => (object)v.ToString()).ToList();

            var cell = commandArgs.First();

            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { valuesToWrite }; //TODO: Fix that it tries to write all data in single row
            try
            {
                var appendRequest = _service.Spreadsheets.Values.Update(valueRange, _defaultSpreadsheetID, cell);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponce = appendRequest.Execute();
                return new CommandExecutionResult($"Spreadsheet updated");
            }
            catch ( Exception ex )
            {
                return new CommandExecutionResult($"ReadValue method exception thrown: {ex.Message}");
            }
        }

        private static bool ValidateCell(object parameter, out CommandExecutionResult errorMessage)
        {
            if (parameter == null || parameter.GetType() == typeof(object))
            {
                errorMessage = new CommandExecutionResult("Usage: /read {cell_from:cell_to} for fetch multiple values or /read {cell} for read single value");
                return false;
            }
            if (parameter is not string cell)
            {
                errorMessage = new CommandExecutionResult("given object is not string. How TF did you make this happen, you bastard???");
                return false;
            }

            var maybeRange = cell.Split(':');
            string regexValidationMatch;

            if (maybeRange.Length == 1)
                regexValidationMatch = @"[A-z][0-9]+";

            else if (maybeRange.Length == 2)
                regexValidationMatch = @"[A-z][0-9]+:[A-z][0-9]+";

            else
            {
                errorMessage = new CommandExecutionResult($"Invalid cell format: {cell}");
                return false;
            }

            if (!Regex.Match(cell.ToString(), regexValidationMatch).Success)
            {
                errorMessage = new CommandExecutionResult($"Invalid cell value given: {cell}");
                return false;
            }
            errorMessage = null;
            return true;
        }

        private static string TableFetchedRangeToString(IList<IList<object>> data, char columnFrom, int rowFrom)
        {
            var result = new StringBuilder();
            var column = columnFrom;

            foreach (var outerList in data)
            {
                foreach (var value in outerList)
                {
                    result.Append($"[{column++}{rowFrom} :: {value}] ");
                }
                rowFrom++;
                column = columnFrom;
                result.Append("\n");
            }

            return result.ToString();
        }
    }
}
