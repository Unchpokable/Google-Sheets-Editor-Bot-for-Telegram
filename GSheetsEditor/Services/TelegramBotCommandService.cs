using GSheetsEditor.Commands;
using GSheetsEditor.Commands.Modules;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
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
            _commandsService.BindModule(typeof(GoogleAPICommands));

            _client.StartReceiving(HandleUpdateAsync, HandlePollingErrorAsync, recieverOptions, cancellationToken: cts.Token);
            Console.WriteLine("Bot ready");
        }

        private TelegramBotClient _client;

        private CommandExecutionBinder _commandsService;

        private async Task HandleUpdateAsync(ITelegramBotClient client, Update update, CancellationToken ct)
        {
            if (update.Message is not { } message)
                return;
            if (message.Text is not { } messageText)
                return;

            var chatID = message.Chat.Id;

            var commandTokens = message.Text.Split(' '); 
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

            var executionResult = await _commandsService.ExecuteAsync(commandTokens[0], commandArgument);

            //await client.SendTextMessageAsync(chatID, executionResult?.Result?.ToString());
            await RouteReply(chatID, client, executionResult);
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
                await client.SendTextMessageAsync(chatID, "Command has not returned reply");

            try
            {
                if (commandResult.ResultType == typeof(Uri))
                {
                    await using Stream stream = new FileStream(((Uri)commandResult.Result).LocalPath, FileMode.Open, FileAccess.Read);
                    await client.SendDocumentAsync(chatID, new InputOnlineFile(stream, $"table_tableid.xlsx"));
                    Task.Run(async () => 
                    {
                        await Task.Delay(3000);
                        System.IO.File.Delete(((Uri)commandResult.Result).LocalPath);
                    });
                }
                else 
                    await client.SendTextMessageAsync(chatID, commandResult.Result.ToString());
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
