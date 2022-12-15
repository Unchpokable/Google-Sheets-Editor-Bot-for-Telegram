using GSheetsEditor.Commands.Assembly;

namespace GSheetsEditor.Commands.Modules
{
    [CommandModule]
    internal class GoogleAPICommands
    {
        [Command("/test")]
        public static CommandExecutionResult TestCommand(object parameter)
        {
            return new CommandExecutionResult(parameter);
        }

        [Command("/multiargs")]
        public static CommandExecutionResult TestMusltiargumentCommand(object parameter)
        {
            return new CommandExecutionResult("Accepted parameters: " + string.Join(" ", (IList<string>)parameter));
        }
    }
}
