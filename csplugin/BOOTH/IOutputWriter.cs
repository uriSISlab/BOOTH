using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    interface IOutputWriter
    {
        void WriteLineArr(string[] line);

        void WriteLine(params string[] line);

        long GetRowNum();

        void Done();
    }
}
