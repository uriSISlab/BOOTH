using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    public interface IOutputWriter
    {
        void WriteLineArr(string[] line, FieldType[] fieldTypes = null);

        void WriteLine(params string[] line);

        long GetRowNum();

        void Done();
    }
}
