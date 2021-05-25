using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    class FileReader : IInputReader
    {

        private readonly StreamReader reader;
        private bool begun;
        private int skipLines;
        private bool endReached;

        public FileReader(string filePath)
        {
            this.reader = new StreamReader(filePath);
            this.begun = false;
            this.skipLines = 0;
            this.endReached = false;
        }

        public bool NoMoreLines()
        {
            if (endReached)
            {
                return true;                
            } else
            {
                if (this.reader.EndOfStream)
                {
                    this.endReached = true;
                    this.reader.Close();
                }
                return this.endReached;
            }
        }

        public string ReadLine()
        {
            if (!this.begun)
            {
                while (skipLines > 0)
                {
                    this.reader.ReadLine();
                    skipLines--;
                }
                begun = true;
            }
            return this.reader.ReadLine();
        }

        public void SetSkipLines(int skipCount)
        {
            this.skipLines = skipCount;
        }
    }
}
