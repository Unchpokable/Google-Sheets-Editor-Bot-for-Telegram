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

            await client.SendTextMessageAsync(chatID, executionResult?.Result?.ToString());
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
    }
}
