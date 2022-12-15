using System.Reflection;

namespace GSheetsEditor.Commands
{
    public class CommandMethodInfo
    {
        public CommandMethodInfo(string commandName, MethodInfo methodInfo)
        {
            CommandName = commandName;
            MethodInfo = methodInfo;
        }

        public string CommandName { get; set; }
        public MethodInfo MethodInfo { get; set; }
    }
    
}
