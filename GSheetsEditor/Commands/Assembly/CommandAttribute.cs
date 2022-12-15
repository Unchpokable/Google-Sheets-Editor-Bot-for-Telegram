namespace GSheetsEditor.Commands.Assembly
{
    internal class CommandAttribute : Attribute
    {
        public CommandAttribute(string command) 
        { 
            Command = command;
        }

        public string Command { get; init; }
    }
}
