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
