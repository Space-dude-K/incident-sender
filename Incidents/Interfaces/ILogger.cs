namespace Incidents.Interfaces
{
    public interface ILogger
    {
        void LogMessage(string msg, bool isEndLine = false);
    }
}
