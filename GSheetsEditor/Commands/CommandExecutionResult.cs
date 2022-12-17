using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot.Types.InputFiles;

namespace GSheetsEditor.Commands
{
    internal class CommandExecutionResult
    {
        public CommandExecutionResult(object result)
        {
            if (result == null)
            {
                IsSuccess = false;
                return;
            }

            ResultType = result.GetType();

            if (!(ResultType == typeof(string) || ResultType == typeof(Uri)))
                throw new ArgumentException("Command return type should be a <string> for sending a text reply or <Uri> for sending files.");

            Result = result;
        }

        public object Result { get; init; }
        public Type ResultType { get; init; }
        public bool IsSuccess { get; init; }
    }
}
