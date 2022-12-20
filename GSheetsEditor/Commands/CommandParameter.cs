using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
