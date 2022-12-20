namespace GSheetsEditor.Commands
{
    internal class Command
    {
        public Command(Func<CommandParameter, CommandExecutionResult> action)
        {
            _execute = action;
        }

        private Func<CommandParameter, CommandExecutionResult> _execute;

        public CommandExecutionResult Execute(CommandParameter arg, bool supressExceptionThrows = false)
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
