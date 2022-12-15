using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

            Result = result;
            ResultType = result.GetType();
        }

        public object Result { get; init; }
        public Type ResultType { get; init; }
        public bool IsSuccess { get; init; }
    }
}
