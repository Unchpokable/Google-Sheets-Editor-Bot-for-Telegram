using GSheetsEditor.Commands.Attributes;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using System.Text.RegularExpressions;
using System.Text;
using GSheetsEditor.Extensions;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Telegram.Bot.Types.ReplyMarkups;
using Telegram.Bot.Types;
using GSheetsEditor.Commands.Modules.Service;
using System.Xml;
using System.Reflection;
using System.Linq;


/*
  __          __     _____  _   _ _____ _   _  _____ 
 \ \        / /\   |  __ \| \ | |_   _| \ | |/ ____|
  \ \  /\  / /  \  | |__) |  \| | | | |  \| | |  __ 
   \ \/  \/ / /\ \ |  _  /| . ` | | | | . ` | | |_ |
    \  /\  / ____ \| | \ \| |\  |_| |_| |\  | |__| |
     \/  \/_/    \_\_|  \_\_| \_|_____|_| \_|\_____|

    !! ACHTUNG !! 
    The code below written with LITERALLY 1 TARGET - IT SHOULD WORK. DOESN'T MATTER HOW, JUST WORK. I'LL TRY TO REFACTOR IT LATER, 
 */
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
        //private static Dictionary<long, SpreadsheetsCollection> _boundSpreadsheets;
        private static BoundSpreadsheetDictionaryProxy _boundSpreadsheets;

        static GoogleAPICommands()
        {
            _boundSpreadsheets = new BoundSpreadsheetDictionaryProxy();
            try
            {
                using (var reader = XmlReader.Create("state_packed.xml"))
                    _boundSpreadsheets.ReadXml(reader);
            }
            catch { }

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

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            var spreadsheetID = _boundSpreadsheets[arg.UserID].Current.ID;

            var parameter = arg?.CommandArgs;

            if (parameter == null || parameter.GetType() == typeof(object))
            {
                return new CommandExecutionResult("Usage: /read {cell_from:cell_to} for fetch multiple values or /read {cell} for read single value");
            }

            if (ValidateCell(parameter, out CommandExecutionResult errorMessage) == CellSelectionType.Invalid)
                return errorMessage;

            var cell = parameter.ToString();

            try
            {
                var userCurrentSpreadsheet = _boundSpreadsheets[arg.UserID].Current;

                var value = _sheetsService.Spreadsheets.Values.Get(spreadsheetID, $"{userCurrentSpreadsheet.WorkingSheet}!{cell}").Execute().Values;
                if (value == null)
                    return new CommandExecutionResult("Empty Cell");
                return new CommandExecutionResult(TableFetchedRangeToString(value, cell[0], int.Parse(cell.From(1).To(':'))));
            }
            catch (Exception ex)
            {
                return new CommandExecutionResult($"Google API method exception thrown: {ex.Message}");
            }
        }

        [Command("/set")]
        public static CommandExecutionResult SetData(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            var cArgs = arg.CommandArgs;

            if (cArgs is not IList<string> commandArgs)
            {
                return new CommandExecutionResult("Invalid Argument. Usage: /set {range || single_cell} {value_to_set}");
            }

            if (commandArgs.Count == 0)
            {
                return new CommandExecutionResult("Invalid argument. Usage: /set {range} {single_value <optional>} for write single value in specified cell. If value not given, sets cells to null");
            }

            var cellSelectionType = ValidateCell(commandArgs[0], out CommandExecutionResult errorMessage);

            if (cellSelectionType == CellSelectionType.Invalid)
                return errorMessage;

            var selectionAmount = CalculateSelectionAmount(commandArgs[0]);

            var valuesRange = PrepareInputToWrite(new List<string>() { "dummy" }.Concat(commandArgs[1].Repeat(selectionAmount)).ToList(), cellSelectionType, commandArgs[0]);

            var cell = $"{_boundSpreadsheets[arg.UserID].Current.WorkingSheet}!{commandArgs.First()}";

            try
            {
                var appendRequest = _sheetsService.Spreadsheets.Values.Update(valuesRange, _boundSpreadsheets[arg.UserID].Current.ID, cell);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponce = appendRequest.Execute();
                return new CommandExecutionResult($"Spreadsheet updated");
            }
            catch (Exception ex)
            {
                return new CommandExecutionResult($"Google API method exception thrown: {ex.Message}");
            }
        }

        [Command("/clear")]
        public static CommandExecutionResult ClearData(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            if (arg.CommandArgs is not string commandArgs)
                return new CommandExecutionResult("/clear command can not take arguments");

            var cellSelectionType = ValidateCell(commandArgs, out CommandExecutionResult errorMessage);

            if (cellSelectionType == CellSelectionType.Invalid)
                return errorMessage;

            var selectionAmount = CalculateSelectionAmount(commandArgs);

            var valuesRange = PrepareInputToWrite(new List<string>() { "dummy" }.Concat("".Repeat(selectionAmount)).ToList(), cellSelectionType, commandArgs);

            var cell = $"{_boundSpreadsheets[arg.UserID].Current.WorkingSheet}!{commandArgs}";

            try
            {
                var appendRequest = _sheetsService.Spreadsheets.Values.Update(valuesRange, _boundSpreadsheets[arg.UserID].Current.ID, cell);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponce = appendRequest.Execute();
                return new CommandExecutionResult($"Spreadsheet updated");
            }
            catch (Exception ex)
            {
                return new CommandExecutionResult($"Google API method exception thrown: {ex.Message}");
            }
        }

        private static int CalculateSelectionAmount(string selectionQuerry)
        {

            if (selectionQuerry.Contains(':'))
            {
                var bounds = selectionQuerry.Split(':');
                if (bounds.Length == 2)
                    return new SelectionToken(bounds[0]).DistanceTo(new SelectionToken(bounds[1]));
            }
            return 1;
        }

        [Command("/write")]
        public static CommandExecutionResult WriteData(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            var spreadsheetID = _boundSpreadsheets[arg.UserID].Current.ID;

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

            var cellSelectionType = ValidateCell(commandArgs[0], out CommandExecutionResult errorMessage);

            if (cellSelectionType == CellSelectionType.Invalid)
                return errorMessage;

            var valueRange = PrepareInputToWrite(commandArgs, cellSelectionType, commandArgs[0]);

            var cell = $"{_boundSpreadsheets[arg.UserID].Current.WorkingSheet}!{commandArgs.First()}";
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

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            var spreadsheetID = _boundSpreadsheets[arg.UserID].Current.ID;

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
                throw new ArgumentNullException(nameof(arg), "Caller API sends a null command argument");

            var pars = arg.CommandArgs;

            try
            {
                var spreadsheet = new Spreadsheet() { Properties = new SpreadsheetProperties() };

                var title = FormatCommandArgsToSpreadsheetTitle(pars);

                spreadsheet.Properties.Title = title;
                var createdSpreadsheet = _sheetsService.Spreadsheets.Create(spreadsheet).Execute();

                var accessPermissions = new Permission() { Type = "anyone", Role = "reader" };
                _driveService.Permissions.Create(accessPermissions, createdSpreadsheet.SpreadsheetId).Execute();

                if (!_boundSpreadsheets.ContainsKey(arg.UserID))
                    _boundSpreadsheets.Add(arg.UserID, new SpreadsheetsCollection());

                _boundSpreadsheets[arg.UserID].Add(new GoogleSpreadsheetInfo(createdSpreadsheet));
                if (_boundSpreadsheets.Keys.Count == 1) //First table for this user - auto switch to it 
                    _boundSpreadsheets[arg.UserID].SwitchTo(0);
                Save();
                return new CommandExecutionResult($"Created spreadsheet: {createdSpreadsheet.Properties.Title}, with ID: {createdSpreadsheet.SpreadsheetId}");
            }
            catch (Exception e)
            {
                return new CommandExecutionResult($"Error during command execution: {e.Message}");
            }
        }

        [Command("/addsheet")]
        public static CommandExecutionResult AddNewSheet(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends a null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            var commandArgs = arg.CommandArgs;

            try
            {
                var title = FormatCommandArgsToSpreadsheetTitle(commandArgs);

                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest();
                var addSheetRequest = new AddSheetRequest() { Properties = new SheetProperties() { Title = title } };

                batchUpdateRequest.Requests = new List<Request>() { new Request { AddSheet = addSheetRequest } };

                _sheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, _boundSpreadsheets[arg.UserID].Current.ID).Execute();

                _boundSpreadsheets[arg.UserID].Current.Sheets.Add(title);
                Save();
                return new CommandExecutionResult($"Successfully updated <Add sheet {title}>");

            }
            catch (Exception e)
            {
                return new CommandExecutionResult($"Error during command execution: {e.Message}");
            }
        }

        [Command("/tables")]
        public static CommandExecutionResult ShowAllUserSpreadsheets(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            return new CommandExecutionResult(string.Join('\n', _boundSpreadsheets[arg.UserID].Select(sheet => sheet.Title ?? "Untitled spreadsheet")));
        }

        [Command("/sheets")]
        public static CommandExecutionResult ShowWorkingSpreadsheetSheets(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            return new CommandExecutionResult(string.Join('\n', _boundSpreadsheets[arg.UserID].Current.Sheets));
        }

        [Command("/switch")]
        public static CommandExecutionResult SwitchUserSpreadsheet(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            var pars = arg.CommandArgs;

            if (pars.GetType() == typeof(object) || pars == null) // User calls /switch without arguments
            {
                var sheets = _boundSpreadsheets[arg.UserID].Select(sheet => sheet.Title ?? "Untitled sheet").ToList();

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
                    _boundSpreadsheets[arg.UserID].SwitchTo(selection);
                }
                catch (ArgumentOutOfRangeException e)
                {
                    return new CommandExecutionResult($"Index {selection} out of range");
                }
                return new CommandExecutionResult($"Switched to {_boundSpreadsheets[arg.UserID].Current?.Title}");
            }

            return new CommandExecutionResult("Invalid argument for /switch command. Please, type /switch without arguments to get selection keyboard or /switch {index} to manually switch to spreadsheet you want");
        }

        [Command("/switchsheet")]
        public static CommandExecutionResult SwitchCurrentWorkingSheet(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            var pars = arg.CommandArgs;

            if (pars.GetType() == typeof(object) || pars == null)
            {
                var sheets = _boundSpreadsheets[arg.UserID].Current.Sheets.Select(s => s ?? "Untitled sheet").ToList();

                var count = 0;
                var inlineKeyboardButtons = new List<List<InlineKeyboardButton>>();

                foreach (var sheet in sheets)
                {
                    inlineKeyboardButtons.Add(
                        new List<InlineKeyboardButton>() { InlineKeyboardButton.WithCallbackData($"{count + 1} : {sheet}", $"/switchsheet {count}") });
                    count++;
                }
                return new CommandExecutionResult("Please, select sheet to switch onto") { ReplyMarkup = new InlineKeyboardMarkup(inlineKeyboardButtons) };
            }

            if (int.TryParse(pars.ToString(), out int selection))
            {
                try
                {
                    _boundSpreadsheets[arg.UserID].Current.ChangeWorkingSheet(selection);
                }
                catch (ArgumentOutOfRangeException e)
                {
                    return new CommandExecutionResult($"Index {selection} is out of range");
                }
                return new CommandExecutionResult($"Switched to {_boundSpreadsheets[arg.UserID].Current.WorkingSheet}");
            }

            return new CommandExecutionResult("Invalid argument for /switchsheet command. Please, type /switchsheet without arguments to get selection keyboard or /switch {index} to manually switch to sheet you want");
        }

        [Command("/url")]
        public static CommandExecutionResult GetSpreadsheetUrl(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");
            try
            {
                var cloudSpreadsheet = _sheetsService.Spreadsheets.Get(_boundSpreadsheets[arg.UserID].Current.ID).Execute();
                return new CommandExecutionResult(cloudSpreadsheet.SpreadsheetUrl);
            }
            catch
            {
                return new CommandExecutionResult("Google drive exception");
            }
        }

        [Command("/bind")]
        public static CommandExecutionResult BindExternalSpreadsheet(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends an null command argument");
            return new CommandExecutionResult("Not implemented");
        }

        [Command("/workspace")]
        public static CommandExecutionResult FormatCurrentWorkspaceReport(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends a null command argument");

            var sb = new StringBuilder();

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            foreach (var spreadsheet in _boundSpreadsheets[arg.UserID])
            {
                sb.Append("== Spreadsheet ==\n");
                sb.Append($"Title: {spreadsheet.Title}");
                sb.Append('\n');
                sb.Append($"Sheets on spreadsheet:\n{string.Join('\n', spreadsheet.Sheets.Select(s => $"- {s}"))}\n\n");
            }
            return new CommandExecutionResult(sb.ToString() ?? "No spreadsheets found for user");
        }

        [Command("/current")]
        public static CommandExecutionResult FormatCurrentSelectedSpreadsheetReport(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends a null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}");

            return new CommandExecutionResult($"Selected spreadsheet: {_boundSpreadsheets[arg.UserID].Current.Title}\n Active sheet: {_boundSpreadsheets[arg.UserID].Current.WorkingSheet}");
        }

        [Command("/synchronize")]
        public static CommandExecutionResult SyncronizeLocalDataWithGoogleDrive(CommandParameter arg)
        {
            if (arg == null)
                throw new ArgumentNullException(nameof(arg), "Caller API sends a null command argument");

            if (_boundSpreadsheets.Keys.Count == 0 || !_boundSpreadsheets.ContainsKey(arg.UserID))
                return new CommandExecutionResult($"No sheets bound to user {arg.UserID}. Nothing to sync");


            var cloudCollection = new SpreadsheetsCollection();

            foreach (var spreadsheet in _boundSpreadsheets[arg.UserID])
            {
                try
                {
                    var cloudSpreadsheet = _sheetsService.Spreadsheets.Get(spreadsheet.ID).Execute();
                    if (cloudSpreadsheet != null)
                    {
                        cloudCollection.Add(new GoogleSpreadsheetInfo(cloudSpreadsheet));
                    }
                }
                catch { }
            }
            _boundSpreadsheets[arg.UserID] = cloudCollection;
            return new CommandExecutionResult("Local data updated with remote origin");
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

        [Command("/save")]
        public static CommandExecutionResult Save(CommandParameter arg)
        {
            if (Save())
                return new CommandExecutionResult("Saved");
            else
                return new CommandExecutionResult("Internal server error. Unable to save");
        }
#endif

        private static bool Save()
        {
            try
            {
                using var writer = XmlWriter.Create("state_packed.xml");
                _boundSpreadsheets.WriteXml(writer);
                return true;
            }
            catch
            { return false; }  
        }

        private static ValueRange PrepareInputToWrite(IList<string> commandArgs, CellSelectionType cellSelectionType, string cell)
        {
            // God give me strength to understand what here happening later

            var valuesToWrite = cellSelectionType == CellSelectionType.Range
                ? commandArgs.Skip(1).Select(v => (object)v.ToString()).ToList()
                : new List<object> { string.Join(' ', commandArgs.Skip(1).Select(v => (object)v.ToString())) };

            if (cellSelectionType == CellSelectionType.Single)
                return new ValueRange() { Values = new List<IList<object>> { valuesToWrite } };


            var formattedValues = new List<List<object>>();

            var line = new List<object>();
            var lineContainer = new List<IList<object>>();

            var cellRange = cell.Split(":");
            if (cellRange.Length > 1)
            {
                var from = new SelectionToken(cellRange[0]);
                var to = new SelectionToken(cellRange[1]);
                var lengthByColumns = from.DistanceColumns(to);
                var lengthByRows = from.DistanceRows(to);
                var insertInColumnCount = 0;
                var insertInRowsCount = 0;

                foreach (var value in valuesToWrite)
                {
                    line.Add(value.ToString());
                    insertInColumnCount++;
                    if (insertInColumnCount > lengthByColumns)
                    {
                        lineContainer.Add(new List<object>(line));
                        line = new List<object>();
                        insertInColumnCount = 0;
                        insertInRowsCount++;

                        if (insertInRowsCount > lengthByRows)
                            break;
                    }
                }
            }
            var valueRange = new ValueRange
            {
                Values = lineContainer
            };
            return valueRange;
        }

        private static CellSelectionType ValidateCell(object parameter, out CommandExecutionResult errorMessage)
        {
            if (parameter is not string cell)
            {
                errorMessage = new CommandExecutionResult("given object is not string");
                return CellSelectionType.Invalid;
            }

            var maybeRange = cell.Split(':');
            string regexValidationMatch;
            CellSelectionType selectionType;

            if (maybeRange.Length == 1)
            {
                regexValidationMatch = @"^[A-z][0-9]+$";
                selectionType = CellSelectionType.Single;
            }

            else if (maybeRange.Length == 2)
            {
                regexValidationMatch = @"^[A-z][0-9]+:[A-z][0-9]+$";
                selectionType = !maybeRange[0].Equals(maybeRange[1]) ? CellSelectionType.Range : CellSelectionType.Single;
            }
            else
            {
                errorMessage = new CommandExecutionResult($"Invalid cell format: {cell}");
                return CellSelectionType.Invalid;
            }

            if (!Regex.Match(cell.ToString(), regexValidationMatch).Success)
            {
                errorMessage = new CommandExecutionResult($"Invalid cell value given: {cell}");
                return CellSelectionType.Invalid;
            }
            errorMessage = null;
            return selectionType;
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

        private static string FormatCommandArgsToSpreadsheetTitle(object commandArgs)
        {
            string title;
            if (commandArgs.GetType() == typeof(string))
            {
                title = commandArgs.ToString();
            }
            else if (commandArgs.GetType() == typeof(IList<string>))
            {
                title = string.Join(" ", (IList<string>)commandArgs);
            }
            else
                title = "New Sheet";
            return title;
        }

        private struct SelectionToken
        {
            public char Column { get; init; }
            public int Row { get; set; }

            public SelectionToken(string source)
            {
                if (source.Length < 2) throw new ArgumentException("Can not create selection token from 1-char string");

                Column = source[0];
                Row = int.Parse(string.Join("", source.Skip(1)));
            }

            public int DistanceTo(SelectionToken other)
            {
                return (other.Column - Column + 1) * (other.Row - Row + 1); //A - B is 1, but there is actually 2 cells so +1;
            }

            public int DistanceRows(SelectionToken other) => other.Row - Row;

            public int DistanceColumns(SelectionToken other) => other.Column - Column;
        }
    }
}
