namespace GSheetsEditor.Commands
{
    internal class CommandParameter
    {
        public CommandParameter(long userID, object commandArgs)
        {
            UserID = userID;
            CommandArgs = commandArgs;
        }

        public long UserID { get; init; }
        public object CommandArgs { get; init; }

        /// <summary>
        /// Gets local file path to file attached to message 
        /// </summary>
        public string? AttachedFile { get; set; }
    }
}
