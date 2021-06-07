namespace BOOTH
{
    public interface IInputReader
    {
        void SetSkipLines(int skipCount);
        bool NoMoreLines();
        string ReadLine();
    }
}
