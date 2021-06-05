using BOOTH.LogProcessors;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors
{
    public interface ILogProcessor
    {
        void SetWriter(IOutputWriter writer);

        void SetFileName(string fileName);

        void ReadLine(string line);

        void WriteHeader();

        bool IsThisLog(Worksheet sheet);

        string GetSeparator();

        void Done();

        string GetUniqueTag();
    }
}
