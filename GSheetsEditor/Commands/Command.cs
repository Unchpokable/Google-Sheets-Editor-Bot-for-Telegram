using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GSheetsEditor.Commands
{
    internal class Command
    {
        public Command(Func<object, CommandExecutionResult> action)
        {
            _execute = action;
        }

        private Func<object, CommandExecutionResult> _execute;

        public CommandExecutionResult Execute(object arg, bool supressExceptionThrows = false)
        {
            try
            {
                return _execute(arg);
            }
            catch (Exception e)
            {
                if (supressExceptionThrows) return new CommandExecutionResult($"Exception during execution command: {e.Message}");

                throw;
            }
        }
    }
}
