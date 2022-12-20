namespace GSheetsEditor.Commands.Attributes
{
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    internal class CommandAttribute : Attribute
    {
        public CommandAttribute(string command) 
        { 
            Command = command;
        }

        public string Command { get; init; }
    }
}
