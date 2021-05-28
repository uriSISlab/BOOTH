﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;

namespace BOOTH
{
    class FileWriter : IOutputWriter
    {

        private readonly StreamWriter stream;
        private readonly CsvWriter csv;
        private long rowNum;

        public FileWriter(string filePath)
        {
            this.stream = new StreamWriter(filePath);
            this.csv = new CsvWriter(this.stream, System.Globalization.CultureInfo.InvariantCulture);
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
            foreach (string field in line)
            {
                csv.WriteField(field);
            }
            csv.NextRecord();
            rowNum++;
        }
    }
}
