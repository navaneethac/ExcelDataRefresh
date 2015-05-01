using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataRefresh
{
    class Util
    {
        /// <summary>
        /// SharePoint URLs have http/https
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool IsLocal(string path)
        {
            if (path.Contains("http"))
                return false;
            else
                return true;
        }

        public static string EncodePath(string path)
        {
            if (IsLocal(path))
            {
                return Path.GetFullPath(path);
            }
            else
            {
                // Look for these special strings to find if the user given input
                // URL is already encoded                
                if (path.Contains("%20") || path.Contains("%2B") || path.Contains("%3D"))
                    return path;
                else
                    return Uri.EscapeUriString(path);
            }
        }
    }
}
