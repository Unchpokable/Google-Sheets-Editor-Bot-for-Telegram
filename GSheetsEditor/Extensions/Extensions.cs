using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace GSheetsEditor.Extensions
{
    public static class Extensions
    {
        public static string From(this string origin, int position) //FlUeNt
        {
            return origin.Substring(position);
        }

        public static string To(this string origin, char stop)
        {
            var sb = new StringBuilder();
            foreach (var c in origin)
            {
                if (c != stop)
                    sb.Append(c);
                else
                    return sb.ToString();
            }
            return sb.ToString();
        }
    }
}
