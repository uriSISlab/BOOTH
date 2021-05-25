using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    class FileWriter : IOutputWriter
    {

        private readonly StreamWriter stream;
        private long rowNum;

        public FileWriter(string filePath)
        {
            this.stream = new StreamWriter(filePath);
            this.rowNum = 1;
        }

        public void Done()
        {
            this.stream.Close();
        }

        public long GetRowNum()
        {
            return rowNum;
        }

        public void SetFieldTypes(FieldType[] types)
        {
        }

        public void WriteLine(params string[] line)
        {
            this.WriteLineArr(line);
        }

        public void WriteLineArr(string[] line, FieldType[] fieldTypes = null)
        {
            // TODO wrap item in quotes if it contains a comma so that output is proper CSV
            this.stream.WriteLine(String.Join(",", line));
            rowNum++;
        }
    }
}
