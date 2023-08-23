using Incidents.Interfaces;
using System.Linq;

namespace Incidents
{
    public class PathManager : IPathManager
    {
        public readonly DateManager dateManager;

        public PathManager(DateManager dateManager)
        {
            this.dateManager = dateManager;
        }

        public ConnectionManagerArchiveCollection ArchivesCollection
        {
            get { return dateManager.excelData.Archives; }
        }
        public ConnectionManagerTemplateCollection TemplatesCollection
        {
            get { return dateManager.excelData.Templates; }
        }
        public ConnectionManagerPathsCollection RegionPathCollection
        {
            get { return dateManager.excelData.RegionPaths; }
        }
        public ConnectionManagerEmailCollection EmailsCollection
        {
            get { return dateManager.excelData.Emails; }
        }

        /// <summary>
        /// Email 1
        /// </summary>
        public string Email1
        {
            get { return EmailsCollection.Cast<ConnectionManagerEmailElement>().ElementAt(0).Email; }
        }
        /// <summary>
        /// Email 2
        /// </summary>
        public string Email2
        {
            get { return EmailsCollection.Cast<ConnectionManagerEmailElement>().ElementAt(1).Email; }
        }
        /// <summary>
        /// Report template path (config)
        /// </summary>
        public string Template1Path
        {
            get { return TemplatesCollection.Cast<ConnectionManagerTemplateElement>().ElementAt(0).Path; }
        }
        /// <summary>
        /// Log template path (config)
        /// </summary>
        public string Template2Path
        {
            get { return TemplatesCollection.Cast<ConnectionManagerTemplateElement>().ElementAt(1).Path; }
        }
        /// <summary>
        /// Archive path (config)
        /// </summary>
        public string ArchivePath
        {
            get { return ArchivesCollection.Cast<ConnectionManagerArchiveElement>().ElementAt(0).Path; }
        }
        /// <summary>
        /// Logger path (config)
        /// </summary>
        public string LoggerPath
        {
            get { return ArchivesCollection.Cast<ConnectionManagerArchiveElement>().ElementAt(1).Path; }
        }
        /// <summary>
        /// Daily log path (.xlsx)
        /// </summary>
        public string AggregateDailyArchiveLogFileName
        {
            get
            {
                return LoggerPath
                    + dateManager.GetPreviousAndWorkingDate().Year
                    + @"\" + dateManager.GetPreviousAndWorkingDate().Month
                    + @"\" + "Archive_"
                    + dateManager.GetPreviousAndWorkingDate().ToString(@"yyyy-MM-dd") + ".xlsx";
            }
        }
        /// <summary>
        /// Daily aggregated file path (.xlsx)
        /// </summary>
        public string AggregateDailyArchiveFileName
        {
            get
            {
                return ArchivePath
                    + "Aggregated" + @"\"
                    + dateManager.GetPreviousAndWorkingDate().Year.ToString()
                    + @"\" + dateManager.GetPreviousAndWorkingDate().Month.ToString()
                    + @"\" + "Журнал инцидентов_"
                    + dateManager.GetPreviousAndWorkingDate().ToString(@"yyyy-MM-dd") + ".xlsx";
            }
        }
        /// <summary>
        /// Main log path (.txt)
        /// </summary>
        public string MainLoggerPath
        {
            get
            {
                return ArchivesCollection.Cast<ConnectionManagerArchiveElement>().ElementAt(1).Path
                    + @"\" + "MainLog_" + dateManager.GetPreviousAndWorkingDate().Year.ToString() + ".txt";
            }
        }
        /// <summary>
        /// Additional log path (.txt)
        /// </summary>
        public string AdditionalLoggerPath
        {
            get
            {
                return ArchivesCollection.Cast<ConnectionManagerArchiveElement>().ElementAt(2).Path
                    + "IncidentLog_" + dateManager.GetPreviousAndWorkingDate().Year.ToString() + ".txt";
            }
        }
        /// <summary>
        /// Archive destination folder
        /// </summary>
        public string RawArchivePath
        {
            get
            {
                return ArchivePath + "Raw" + @"\"
                    + dateManager.GetPreviousAndWorkingDate().Year.ToString()
                    + @"\" + dateManager.GetPreviousAndWorkingDate().Month.ToString() + @"\"
                    + dateManager.GetPreviousAndWorkingDate().Day.ToString();
            }
        }
        /// <summary>
        /// Recipient address
        /// </summary>
        public string Email
        {
            get { return EmailsCollection.Cast<ConnectionManagerEmailElement>().ElementAt(0).Email; }
        }
        /// <summary>
        /// Email subject
        /// </summary>
        public string EmailSubject
        {
            get { return "Инциденты Могилёвская область " + dateManager.GetPreviousAndWorkingDate().ToString(@"yyyy.MM.dd"); }
        }
    }
}
