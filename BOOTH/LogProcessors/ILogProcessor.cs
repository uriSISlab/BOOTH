using Microsoft.Office.Interop.Excel;

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
