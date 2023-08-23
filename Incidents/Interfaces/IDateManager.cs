using System;

namespace Incidents.Interfaces
{
    public interface IDateManager
    {
        ConnectionManagerNotWorkingDaysCollection NotWorkingDaysCollection { get; }
        ConnectionManagerWorkingDaysCollection WorkingDaysCollection { get; }

        DateTime DateParser(string rawDate);
        string GetFolderModifiedDate(string folder);
        DateTime GetPreviousAndWorkingDate();
        bool IsWeekEnd(DateTime date);
        bool IsWorkingDay(DateTime date);
        void LoadDates();
        string SetData(string rawStr);
    }
}
