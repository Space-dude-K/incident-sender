using Incidents.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace Incidents
{
    public class StringHelper : IStringHelper
    {
        public readonly ILogger logger;
        private readonly IDateManager dateManager;

        public StringHelper(ILogger logger, IDateManager dateManager)
        {
            this.logger = logger;
            this.dateManager = dateManager;
        }
        #region Проверки
        public Tuple<bool, bool> FileExistAndEmpty(string path)
        {
            if (File.Exists(path))
            {
                if (new FileInfo(path).Length == 0)
                {
                    logger.LogMessage(path + " -> file is empty.");

                    return new Tuple<bool, bool>(false, true);
                }
                else
                {
                    return new Tuple<bool, bool>(true, false);
                }
            }
            else
            {
                return new Tuple<bool, bool>(false, false);
            }
        }
        // Correct type of incident
        public string FixedIncidentType(string str)
        {
            string fixedStr = string.Empty;

            if (str.Contains("Нет") && str.Contains("инц"))
            {
                logger.LogMessage("Fixing incident type -> " + str);
                fixedStr = "Нет инцидентов";
            }

            return fixedStr;
        }
        public bool DateExist(string date, string region)
        {
            Console.WriteLine("[DateExist] COMPARE: for {0} - {1} AND {2}", 
                region, date, dateManager.GetPreviousAndWorkingDate().ToShortDateString());
            logger.LogMessage("[DateExist] COMPARE: for " 
                + region + " - " + date + " AND " + dateManager.GetPreviousAndWorkingDate().ToShortDateString());

            if (date == "Null" || date == "Empty" || date == "Skip")
            {
                return false;
            }
            else if (DateTime.TryParse(date, out DateTime res) && dateManager.GetPreviousAndWorkingDate() == res)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public Color SetCellColor(bool state, int counter, bool skip)
        {
            if (new HashSet<int> { 0, 2, 22, 23 }.Contains(counter) && skip == true)
            {
                return Color.White;
            }
            else
            {
                return state == true ? Color.Green : Color.Red;
            }
        }
        public string SetLogString(int counter, string msg, bool skip)
        {
            if (new HashSet<int> { 0, 2, 22, 23 }.Contains(counter) && skip == false)
            {
                return msg;
            }
            else
            {
                return new HashSet<int> { 0, 2, 22, 23 }.Contains(counter) && skip == true ? "-" : msg;
            }
        }
        public bool IsStringsIsEmptyOrNull(string str1, string str2)
        {
            bool result = (str1 == "Empty" || str2 == "Empty") || (string.IsNullOrEmpty(str1) || string.IsNullOrEmpty(str2)) ? true : false;
            Console.WriteLine("STR1 {0}, STR2 {1}", str1, str2);
            return result;
        }
        #endregion
    }
}