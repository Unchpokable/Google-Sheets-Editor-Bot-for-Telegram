using GSheetsEditor.Commands.Assembly;
using System.Reflection;

namespace GSheetsEditor.Commands
{
    internal partial class CommandExecutionBinder
    {
        public CommandExecutionBinder()
        {
            _bindedCommands = new Dictionary<string, Command>();
        }

        private Dictionary<string, Command> _bindedCommands;

        public void Bind(string command, Func<object, CommandExecutionResult> execute)
        {
            var cmd = new Command(execute);
            _bindedCommands.Add(command, cmd);
        }

        public void Bind(CommandMethodInfo command)
        {
            var commandPrompt = command.CommandName;
            var commandExecutable = (Func<object, CommandExecutionResult>)Delegate.CreateDelegate(typeof(Func<object, CommandExecutionResult>), command.MethodInfo);
            Bind(commandPrompt, commandExecutable);
        }

        public void Unbind(string command)
        {
            _bindedCommands.Remove(command);
        }

        public void BindModule(Type moduleType)
        {
            if (moduleType.GetCustomAttribute<CommandModuleAttribute>() == null)
                throw new ArgumentException($"Given type does not contain {typeof(CommandModuleAttribute)} attribute. Can not exctract commands");

            var potentialCommands = moduleType.GetMethods(BindingFlags.Public | BindingFlags.Static).Where(method => method.GetCustomAttribute<CommandAttribute>() != null)
                .Select(method => new CommandMethodInfo(method.GetCustomAttribute<CommandAttribute>().Command, method))
                .ToList();

            if (potentialCommands.Count() == 0) return;

            var executableCommands = potentialCommands.Where(AssertMethodMatchDefaultCommandSignature).ToList();
            executableCommands.ForEach(Bind);
        }

        public async Task<CommandExecutionResult> ExecuteAsync(string command, object arg)
        {
            if (!_bindedCommands.ContainsKey(command))
                return new CommandExecutionResult($"Command not recognized: {command}");

            return await Task.Run(() => _bindedCommands[command].Execute(arg));
        }

        private bool AssertMethodMatchDefaultCommandSignature(CommandMethodInfo method)
        {
            if (method == null) return false;
            return AssertMethodMatchDefaultCommandSignature(method.MethodInfo);
        }

        private bool AssertMethodMatchDefaultCommandSignature(MethodInfo method)
        {
            if (method == null) return false;

            return method.ReturnType == typeof(CommandExecutionResult) 
                && method.GetParameters().Length == 1
                && method.GetParameters()[0].ParameterType == typeof(object);
        }
    }
}
