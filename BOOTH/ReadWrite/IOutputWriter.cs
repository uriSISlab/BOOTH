using System.Collections.Generic;

namespace BOOTH
{
    public interface IOutputWriter
    {
        void WriteLineArr(IEnumerable<string> line, IEnumerable<FieldType> fieldTypes = null);

        void WriteLine(params string[] line);

        long GetRowNum();

        void Done();
    }
}
