using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOOTH
{
    interface IInputReader
    {
        void SetSkipLines(int skipCount);
        bool NoMoreLines();
        string ReadLine();
    }
}
