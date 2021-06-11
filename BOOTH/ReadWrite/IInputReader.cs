namespace BOOTH
{
    public interface IInputReader
    {
        bool NoMoreLines();
        string ReadLine();
    }
}
