using Incidents.Interfaces;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace Incidents
{
    public class DateManager : IDateManager
    {
        public ExcelDataSection excelData;
        public DateDataSection datesData;
        // Праздники
        private readonly HashSet<DateTime> NotWorkingDays = new HashSet<DateTime>();
        // Переносы
        private readonly HashSet<DateTime> WorkingDays = new HashSet<DateTime>();

        public DateManager()
        {
            // Grab the Environments listed in the App.config and add them to list.
            excelData = ConfigurationManager.GetSection(ExcelDataSection.SectionName) as ExcelDataSection;
            datesData = ConfigurationManager.GetSection(DateDataSection.SectionName) as DateDataSection;
        }

        public ConnectionManagerNotWorkingDaysCollection NotWorkingDaysCollection
        {
            get { return datesData.NotWorkingDays; }
        }
        public ConnectionManagerWorkingDaysCollection WorkingDaysCollection
        {
            get { return datesData.WorkingDays; }
        }
        #region Работа с датами
        /// <summary>
        /// Date initialize
        /// </summary>
        public void LoadDates()
        {
            foreach (ConnectionManagerNotWorkingDaysElement holiday in 
                NotWorkingDaysCollection.Cast<ConnectionManagerNotWorkingDaysElement>())
            {
                var dtH = DateTime.Parse(holiday.Date);
                Console.WriteLine("Not working days from config is loaded. -> {0}", dtH.ToShortDateString());
                NotWorkingDays.Add(dtH);
            }
            foreach (ConnectionManagerWorkingDaysElement workingday in 
                WorkingDaysCollection.Cast<ConnectionManagerWorkingDaysElement>())
            {
                var dtW = DateTime.Parse(workingday.Date);
                Console.WriteLine("Working days from config is loaded. -> {0}", dtW.ToShortDateString());
                WorkingDays.Add(dtW);
            }
        }
        public bool IsWeekEnd(DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        }
        public bool IsWorkingDay(DateTime date)
        {
            return WorkingDays.Contains(date.Date) || !IsWeekEnd(date) && !NotWorkingDays.Contains(date.Date);
        }
        public DateTime GetPreviousAndWorkingDate()
        {
            DateTime prevDate = DateTime.Now.Subtract(new TimeSpan(1, 0, 0, 0)).Date;

            while (!IsWorkingDay(prevDate))
            {
                prevDate = prevDate.Subtract(new TimeSpan(1)).Date;
            }

            return prevDate;
        }
        public DateTime DateParser(string rawDate)
        {
            DateTime dt = GetPreviousAndWorkingDate();
            string parsedValue = rawDate.Substring(0, rawDate.IndexOf(','));

            if (DateTime.TryParse(parsedValue, out dt))
            {
                Console.WriteLine("[DATE PARSER] Ol date parsed: " + dt.ToShortDateString() + ".");
            }
            else
            {
                Console.WriteLine("[DATE PARSER] Ol date parser error for " + parsedValue + ". Using the previous working date.");
            }

            return dt;
        }
        public string GetFolderModifiedDate(string folder)
        {
            string result = string.Empty;

            try
            {
                if (!Directory.Exists(folder))
                {
                    Console.WriteLine("[GetFolderModifiedDate] Directory -> " + folder + " didn't exist.");
                    result = "-";
                }
                else
                {
                    // Get the creation time of a well-known directory.
                    result = Directory.GetLastWriteTime(folder).ToString();
                    Console.WriteLine("The last write time for " + folder + " directory was " + result + ".");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
                result = "Error";
            }

            return result;
        }
        public string SetData(string rawStr)
        {
            if (rawStr.Equals("Empty") || string.IsNullOrEmpty(rawStr) || string.IsNullOrWhiteSpace(rawStr))
            {
                return string.Empty;
            }
            else
            {
                return rawStr;
            }
        }
        #endregion
    }
}