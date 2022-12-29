using GSheetsEditor.Commands;
using GSheetsEditor.Commands.Modules;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.InputFiles;

namespace GSheetsEditor.Services
{
    internal class TelegramBotCommandService : ServiceBase
    {
        public TelegramBotCommandService(IServiceProvider serviceProvider, IConfigurationRoot configuration) : base(serviceProvider, configuration)
        {
            using var cts = new CancellationTokenSource();

            _client = new TelegramBotClient(_configuration["TelegramApiKey"]);

            var recieverOptions = new ReceiverOptions()
            {
                AllowedUpdates = Array.Empty<UpdateType>()
            };

            _commandsService = new CommandExecutionBinder();
            _commandsService.BindAssembly(Assembly.GetExecutingAssembly());

            _client.StartReceiving(HandleUpdateAsync, HandlePollingErrorAsync, recieverOptions, cancellationToken: cts.Token);
#if DEBUG
            Console.WriteLine($"[{DateTime.Now} :: INFO] Bot ready");
#endif
        }

        private TelegramBotClient _client;

        private CommandExecutionBinder _commandsService;

        private async Task HandleUpdateAsync(ITelegramBotClient client, Update update, CancellationToken ct)
        {
            if (update.CallbackQuery != null)
            {
                await ProcessCallbackQuerry(client, update.CallbackQuery, ct);
                return;
            }

            await ProcessCommand(client, update, ct);
        }

        private async Task ProcessCallbackQuerry(ITelegramBotClient client, CallbackQuery callbackQuerry, CancellationToken ct)
        {
            if (callbackQuerry.Data is not string querry)
                return;

            var userID = callbackQuerry.Message.Chat.Id;
            await ProcessCommand(client, userID, callbackQuerry.Data, ct);
        }

        private async Task ProcessCommand(ITelegramBotClient client, Update update, CancellationToken ct)
        {
            // Do not write code at 5 A.M.
            if (update.Message is not { } message)
                return;

            string command = "";

            if (update.Message.Text == null)
            {
                if (update.Message.Caption != null)
                    command = update.Message.Caption;
            }
            else 
                command = update.Message.Text;

            var chatID = message.Chat.Id;

            await ProcessCommand(client, chatID, command, ct, update.Message.Document);
        }

        private async Task ProcessCommand(ITelegramBotClient client, long chatID, string messageText, CancellationToken ct, Document attachedFile = null)
        {
            var commandTokens = messageText.Split(' ');
            if (commandTokens.Length == 0)
            {
                await client.SendTextMessageAsync(chatID, "Empty string. Can not execute");
                return;
            }

            object commandArgument;

            if (commandTokens.Length >= 2)
            {
                commandArgument = commandTokens.Length == 2 ? commandTokens[1] : (object)commandTokens.Skip(1).ToList();
            }
            else
                commandArgument = new object();

            var commandParameter = new CommandParameter(chatID, commandArgument);

            string attachedFileLocalPath = null;

            if (attachedFile != null)
            {
                attachedFileLocalPath = $"{AppContext.BaseDirectory}\\{attachedFile.FileName}";
                await using var fileStream = new FileStream(attachedFileLocalPath, FileMode.Create, FileAccess.Write);
                await client.DownloadFileAsync((await client.GetFileAsync(attachedFile.FileId)).FilePath, fileStream);
                commandParameter.AttachedFile = attachedFileLocalPath;
            }

            var executionResult = await _commandsService.ExecuteAsync(commandTokens[0], commandParameter);
            await RouteReply(chatID, client, executionResult);

            if (attachedFileLocalPath != null)
            {
                System.IO.File.Delete(attachedFileLocalPath);
            }
        }

        private async Task HandlePollingErrorAsync(ITelegramBotClient client, Exception exception, CancellationToken ct)
        {
            var errorMessage = exception switch
            {
                ApiRequestException apiRequestException => $"Telegram API Error:\n{apiRequestException.ErrorCode}\n{apiRequestException.Message}\n\n========================================\n\n",
                _ => exception.ToString()
            };

            Console.Write(errorMessage);
        }

        private async Task RouteReply(ChatId chatID, ITelegramBotClient client, CommandExecutionResult commandResult)
        {
            if (commandResult == null)
            {
                await client.SendTextMessageAsync(chatID, "Command has not returned reply");
                return;
            }
            try
            {
                if (commandResult.ResultType == typeof(Uri))
                {
                    await using Stream stream = new FileStream(((Uri)commandResult.Result).LocalPath, FileMode.Open, FileAccess.Read);
                    await client.SendDocumentAsync(chatID, new InputOnlineFile(stream, $"table_tableid.xlsx"));

#pragma warning disable CS4014 //Ye, ye, VS, not awaited task... I actually know what I'm doing here, ok?
                    Task.Run(async () => 
                    {
                        await Task.Delay(3000);
                        System.IO.File.Delete(((Uri)commandResult.Result).LocalPath);
                    });
#pragma warning restore CS4014
                }
                else 
                    await client.SendTextMessageAsync(chatID, commandResult.Result.ToString(), replyMarkup: commandResult.ReplyMarkup);
            }

            catch (ApiRequestException exception)
            {
                await client.SendTextMessageAsync(chatID, "Command returned value that cause Telegram API error");
            }

            catch (Exception exception) 
            {
                await client.SendTextMessageAsync(chatID, $"Critical System Error: {exception.Message}\nIf you see this message please contact bot support");
            }
        }
    }
}
