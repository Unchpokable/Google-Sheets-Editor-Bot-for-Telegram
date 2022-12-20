using GSheetsEditor.Commands.Attributes;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using System.Text.RegularExpressions;
using System.Text;
using GSheetsEditor.Extensions;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Drive.v3;
using Telegram.Bot.Types.ReplyMarkups;

namespace GSheetsEditor.Commands.Modules
{
    [CommandModule]
    internal class GoogleAPICommands
    {
        private static readonly string[] _scopes = { SheetsService.Scope.Spreadsheets, SheetsService.Scope.Drive, DriveService.Scope.Drive };
        private static readonly string _applicationName = "app";
        private static string _defaultSpreadsheetID = "1wG0yquhycja9AWrf1xjvCbkDW5UA2BTPtQqeyKhzST0";
        private static readonly string _baseSheet = "Sheet1";
        private static SheetsService _sheetsService;
        private static DriveService _driveService;
        private static Dictionary<long, SpreadsheetsCollection> _boundSpreadsheets;

        static GoogleAPICommands()
        {
            _boundSpreadsheets = new Dictionary<long, SpreadsheetsCollection>();

            GoogleCredential credentials;

            using (var file = new FileStream("GoogleApiCredentials.json", FileMode.Open, FileAccess.Read))
            {
                credentials = GoogleCredential.FromStream(file)
                                .CreateScoped(_scopes);
            }

            _sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credentials,
                ApplicationName = _applicationName,
            });

            _driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credentials,
                ApplicationName = _applicationName,
            });
        }

        [Command("/read")]
        public static CommandExecutionResult ReadData(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends null command argument");

            if (!GetActualSpreadsheetIDForUser(arg.UserID, out string spreadsheetID))
            {
                return new CommandExecutionResult($"No spreadsheets bound to user {arg.UserID}. Please, create a new spreadsheet using command /new or bind existing spreadsheet using commnad /bind");
            }

            var parameter = arg?.CommandArgs;

            if (parameter == null || parameter.GetType() == typeof(object))
            {
                return new CommandExecutionResult("Usage: /read {cell_from:cell_to} for fetch multiple values or /read {cell} for read single value");
            }

            if (!ValidateCell(parameter, out CommandExecutionResult errorMessage))
                return errorMessage;

            var cell = parameter.ToString();

            try
            {
                var value = _sheetsService.Spreadsheets.Values.Get(spreadsheetID, $"{_baseSheet}!{cell}").Execute().Values;
                if (value == null)
                    return new CommandExecutionResult("Empty Cell");
                return new CommandExecutionResult(TableFetchedRangeToString(value, cell[0], int.Parse(cell.From(1).To(':'))));
            }
            catch (Exception ex)
            {
                return new CommandExecutionResult($"Google API method exception thrown: {ex.Message}");
            }
        }

        [Command("/select")]
        public static CommandExecutionResult SelectData(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends null command argument");

            return new CommandExecutionResult("Not implemented now");
        }

        [Command("/write")]
        public static CommandExecutionResult WriteData(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            if (!GetActualSpreadsheetIDForUser(arg.UserID, out string spreadsheetID))
            {
                return new CommandExecutionResult($"No spreadsheets bound to user {arg.UserID}. Please, create a new spreadsheet using command /new or bind existing spreadsheet using commnad /bind");
            }

            var parameter = arg?.CommandArgs;

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
                var appendRequest = _sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetID, cell);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponce = appendRequest.Execute();
                return new CommandExecutionResult($"Spreadsheet updated");
            }
            catch (Exception ex)
            {
                return new CommandExecutionResult($"Google API method exception thrown: {ex.Message}");
            }
        }

        [Command("/export")]
        public static CommandExecutionResult ExportSpreadsheet(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            if (!GetActualSpreadsheetIDForUser(arg.UserID, out string spreadsheetID))
            {
                return new CommandExecutionResult($"No spreadsheets bound to user {arg.UserID}. Please, create a new spreadsheet using command /new or bind existing spreadsheet using commnad /bind");
            }

            var parameter = arg?.CommandArgs;

            var request = _driveService.Files.List();
            request.Q = "mimeType = 'application/vnd.google-apps.spreadsheet'";
            var response = request.Execute();

            if (response == null)
                return new CommandExecutionResult("Drive Error: No such spreadsheet");

            var requestedSpreadsheet = response.Files.FirstOrDefault(file => file.Id == spreadsheetID);

            if (requestedSpreadsheet == null)
                return new CommandExecutionResult("Drive Error: No such spreadsheet");

            var localPath = $"{AppContext.BaseDirectory}\\local-{spreadsheetID}.xlsx";


            using (var localFile = new FileStream(localPath, FileMode.Create, FileAccess.ReadWrite))
            {
                var file = _driveService.Files.Export(requestedSpreadsheet.Id, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                file.MediaDownloader.ProgressChanged += progress =>
                {
                    switch (progress.Status)
                    {
                        case Google.Apis.Download.DownloadStatus.Failed:
                            Console.WriteLine($"[ERR] :: Download Failed: {progress.Exception}");
                            break;
                    }
                };

                file.Download(localFile);
                var fileUri = new Uri(localPath, UriKind.Absolute);
                return new CommandExecutionResult(fileUri);
            }

        }

        [Command("/new")]
        public static CommandExecutionResult CreateNewSpreadsheet(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            var pars = arg.CommandArgs;

            if (pars == null)
                return new CommandExecutionResult("Command execution failed: no requiered parameters given.");
            try
            {
                var spreadsheet = new Spreadsheet() { Properties = new SpreadsheetProperties() };
                
                string title;
                if (pars.GetType() == typeof(string))
                {
                    title = pars.ToString();
                }
                else if (pars.GetType() == typeof(List<string>))
                {
                    title = string.Join(" ", (IList<string>)pars);
                }
                else
                    title = $"New Spreadsheet - At {DateTime.Now}";

                spreadsheet.Properties.Title = title;
                var createdSpreadsheet = _sheetsService.Spreadsheets.Create(spreadsheet).Execute();

                if (!_boundSpreadsheets.ContainsKey(arg.UserID))
                    _boundSpreadsheets.Add(arg.UserID, new SpreadsheetsCollection());

                _boundSpreadsheets[arg.UserID].Add(createdSpreadsheet);
                return new CommandExecutionResult($"Created spreadsheet: {createdSpreadsheet.Properties.Title}, with ID: {createdSpreadsheet.SpreadsheetId}");
            }
            catch (Exception e)
            {
                return new CommandExecutionResult($"Error during command execution: {e.Message}");
            }
        }

        [Command("/sheets")]
        public static CommandExecutionResult ShowAllUserSpreadsheets(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            if (_boundSpreadsheets.Keys.Count == 0)
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            return new CommandExecutionResult(string.Join('\n', _boundSpreadsheets[arg.UserID].Select(sheet => sheet.Properties?.Title ?? "Untitled spreadsheet")));
        }

        [Command("/switch")]
        public static CommandExecutionResult SwitchUserSpreadsheet(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            if (_boundSpreadsheets.Keys.Count == 0)
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            var pars = arg.CommandArgs;

            if (pars.GetType() == typeof(object) || pars == null) // User calls /switch without arguments
            {
                var sheets = _boundSpreadsheets[arg.UserID].Select(sheet => sheet.Properties?.Title ?? "Untitled sheet").ToList();

                var count = 1;
                var inlineKeyboardButtons = new List<List<InlineKeyboardButton>>();

                foreach (var sheet in sheets)
                {
                    inlineKeyboardButtons.Add(
                        new List<InlineKeyboardButton>() { InlineKeyboardButton.WithCallbackData($"{count} : {sheet}", $"/switch {count - 1}") });
                    count++;
                }

                return new CommandExecutionResult("Please, select spreadsheet to switch onto") { ReplyMarkup = new InlineKeyboardMarkup(inlineKeyboardButtons) };
            }
            if (int.TryParse(pars.ToString(), out int selection))
            {
                try
                {
                    _boundSpreadsheets[arg.UserID].SwitchTo(int.Parse(pars.ToString()));
                }
                catch (ArgumentOutOfRangeException e)
                {
                    return new CommandExecutionResult($"Index {int.Parse(pars.ToString())} out of range");
                }
                return new CommandExecutionResult($"Switched to {_boundSpreadsheets[arg.UserID].Current.Properties.Title}");
            }

            return new CommandExecutionResult("Invalid argument for /switch command. Please, type /switch without arguments to get selection keyboard or /switch {index} to manually switch to spreadsheet you want");
        }

#if DEBUG
        [Command("/cleardrive")]
        public static CommandExecutionResult ClearDrive(CommandParameter arg)
        {
            var files = _driveService.Files.List();
            files.Q = "mimeType = 'application/vnd.google-apps.spreadsheet'";
            var response = files.Execute();
            var counter = 0;
            var failcounter = 0;

            foreach (var file in response.Files)
            {
                try
                {
                    _driveService.Files.Delete(file.Id).Execute();
                    counter++;
                }
                catch
                {
                    failcounter++;
                }
            }

            return new CommandExecutionResult($"Successfully removed {counter} files. {failcounter} files remove fail");
        }

        [Command("/listdrive")]
        public static CommandExecutionResult ListBotDrive(CommandParameter arg) 
        {
            var files = _driveService.Files.List();
            files.Q = "mimeType = 'application/vnd.google-apps.spreadsheet'";
            var response = files.Execute();

            var result = new StringBuilder();

            foreach (var file in response.Files)
            {
                result.Append($"{file.Id}\n");
            }

            return new CommandExecutionResult(result.ToString());
        } 
#endif
        private static bool GetActualSpreadsheetIDForUser(long uID, out string spreadsheet)
        {
            if (!_boundSpreadsheets.ContainsKey(uID))
            {
#if DEBUG //It's OK when debug build uses some default spreadsheet for tests that i have access to, but its NOT OK for release build
                spreadsheet = _defaultSpreadsheetID;
                return true;
#else
                return false;
#endif
            }
            else
                spreadsheet = _boundSpreadsheets[uID].Current.SpreadsheetId;
            return true;
        }

        private static bool ValidateCell(object parameter, out CommandExecutionResult errorMessage)
        {
            if (parameter is not string cell)
            {
                errorMessage = new CommandExecutionResult("given object is not string. How TF did you make this happen, you bastard???");
                return false;
            }

            var maybeRange = cell.Split(':');
            string regexValidationMatch;

            if (maybeRange.Length == 1)
                regexValidationMatch = @"^[A-z][0-9]+$";

            else if (maybeRange.Length == 2)
                regexValidationMatch = @"^[A-z][0-9]+:[A-z][0-9]+$";

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
                result.Append('\n');
            }

            return result.ToString();
        }
    }
}
