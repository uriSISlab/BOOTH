using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    public enum LogType
    {
        VSAP_BMD,
        DICE,
        DICX,
        UNKNOWN
    }

    public enum IOType
    {
        FILE,
        SHEET
    }
    public static class Util
    {
        public static string getLetterFromNumber(int number)
        {
            return ((char)('A' + number)).ToString();
        }
    }
}
