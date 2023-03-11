using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Reflection;
using Incidents.Interfaces;
using System.Globalization;

namespace Incidents
{
    class DailyReport
    {
        enum ErrorMsgs
        {
            [Description("Отсутствует файл")]
            MissingFile,
            [Description("Неправильная дата")]
            WrongDate,
            [Description("Неправильный формат даты")]
            WrongDateFormat,
            [Description("Отсутствует описание или меры")]
            MissingStrings,
            [Description("Оотсутствует инцидент")]
            MissingIncident,
            [Description("Есть инциденты, но неправильный тип инцидента")]
            WrongIncident,
            [Description("Пустой файл")]
            FileIsEmpty
        }

        private ExcelDataSection excelData;
        private DateDataSection datesData;
        private ILogger logger;

        // Праздники
        private readonly HashSet<DateTime> NotWorkingDays = new HashSet<DateTime>();
        // Переносы
        private readonly HashSet<DateTime> WorkingDays = new HashSet<DateTime>();

        private ConnectionManagerArchiveCollection ArchivesCollection
        {
            get { return excelData.Archives; }
        }
        private ConnectionManagerTemplateCollection TemplatesCollection
        {
            get { return excelData.Templates; }
        }
        private ConnectionManagerPathsCollection RegionPathCollection
        {
            get { return excelData.RegionPaths; }
        }
        private ConnectionManagerEmailCollection EmailsCollection
        {
            get { return excelData.Emails; }
        }
        private ConnectionManagerNotWorkingDaysCollection NotWorkingDaysCollection
        {
            get { return datesData.NotWorkingDays; }
        }
        private ConnectionManagerWorkingDaysCollection WorkingDaysCollection
        {
            get { return datesData.WorkingDays; }
        }
        /// <summary>
        /// Email 1
        /// </summary>
        private string Email1
        {
            get { return EmailsCollection.Cast<ConnectionManagerEmailElement>().ElementAt(0).Email; }
        }
        /// <summary>
        /// Email 2
        /// </summary>
        private string Email2
        {
            get { return EmailsCollection.Cast<ConnectionManagerEmailElement>().ElementAt(1).Email; }
        }
        /// <summary>
        /// Report template path (config)
        /// </summary>
        private string Template1Path
        {
            get { return TemplatesCollection.Cast<ConnectionManagerTemplateElement>().ElementAt(0).Path; }
        }
        /// <summary>
        /// Log template path (config)
        /// </summary>
        private string Template2Path
        {
            get { return TemplatesCollection.Cast<ConnectionManagerTemplateElement>().ElementAt(1).Path; }
        }
        /// <summary>
        /// Archive path (config)
        /// </summary>
        private string ArchivePath
        {
            get { return ArchivesCollection.Cast<ConnectionManagerArchiveElement>().ElementAt(0).Path; }
        }
        /// <summary>
        /// Logger path (config)
        /// </summary>
        private string LoggerPath
        {
            get { return ArchivesCollection.Cast<ConnectionManagerArchiveElement>().ElementAt(1).Path; }
        }
        /// <summary>
        /// Daily log path (.xlsx)
        /// </summary>
        private string AggregateDailyArchiveLogFileName
        {
            get { return LoggerPath + GetPreviousAndWorkingDate().Year + @"\" + GetPreviousAndWorkingDate().Month + @"\" + "Archive_" + GetPreviousAndWorkingDate().ToString(@"yyyy-MM-dd") + ".xlsx"; }
        }
        /// <summary>
        /// Daily aggregated file path (.xlsx)
        /// </summary>
        private string AggregateDailyArchiveFileName
        {
            get { return ArchivePath + "Aggregated" + @"\" + GetPreviousAndWorkingDate().Year.ToString() + @"\" + GetPreviousAndWorkingDate().Month.ToString() + @"\" + "Журнал инцидентов_" + GetPreviousAndWorkingDate().ToString(@"yyyy-MM-dd") + ".xlsx"; }
        }
        /// <summary>
        /// Main log path (.txt)
        /// </summary>
        private string MainLoggerPath
        {
            get { return ArchivesCollection.Cast<ConnectionManagerArchiveElement>().ElementAt(1).Path + @"\" + "MainLog_" + GetPreviousAndWorkingDate().Year.ToString() + ".txt"; }
        }
        /// <summary>
        /// Additional log path (.txt)
        /// </summary>
        private string AdditionalLoggerPath
        {
            get { return ArchivesCollection.Cast<ConnectionManagerArchiveElement>().ElementAt(2).Path + "IncidentLog_" + GetPreviousAndWorkingDate().Year.ToString() + ".txt"; }
        }
        /// <summary>
        /// Archive destination folder
        /// </summary>
        private string RawArchivePath
        {
            get { return ArchivePath + "Raw" + @"\" + GetPreviousAndWorkingDate().Year.ToString() + @"\" + GetPreviousAndWorkingDate().Month.ToString() + @"\" + GetPreviousAndWorkingDate().Day.ToString(); }
        }
        /// <summary>
        /// Recipient address
        /// </summary>
        private string Email
        {
            get { return EmailsCollection.Cast<ConnectionManagerEmailElement>().ElementAt(0).Email; }
        }
        /// <summary>
        /// Email subject
        /// </summary>
        private string EmailSubject
        {
            get { return "Инциденты Могилёвская область " + GetPreviousAndWorkingDate().ToString(@"yyyy.MM.dd"); }
        }
        public DailyReport()
        {
            // Grab the Environments listed in the App.config and add them to list.
            excelData = ConfigurationManager.GetSection(ExcelDataSection.SectionName) as ExcelDataSection;
            datesData = ConfigurationManager.GetSection(DateDataSection.SectionName) as DateDataSection;

            // Load dates
            LoadDates();
            Console.WriteLine("Load logger");
            // Load logger
            logger = new LogToTxt(new List<string>() { MainLoggerPath });
            Console.WriteLine("Load complete");
        }
        #region Работа с датами
        /// <summary>
        /// Date initialize
        /// </summary>
        private void LoadDates()
        {
            foreach (ConnectionManagerNotWorkingDaysElement holiday in NotWorkingDaysCollection.Cast<ConnectionManagerNotWorkingDaysElement>())
            {
                var dtH = DateTime.Parse(holiday.Date);
                Console.WriteLine("Not working days from config is loaded. -> {0}", dtH.ToShortDateString());
                NotWorkingDays.Add(dtH);
            }
            foreach (ConnectionManagerWorkingDaysElement workingday in WorkingDaysCollection.Cast<ConnectionManagerWorkingDaysElement>())
            {
                var dtW = DateTime.Parse(workingday.Date);
                Console.WriteLine("Working days from config is loaded. -> {0}", dtW.ToShortDateString());
                WorkingDays.Add(dtW);
            }
        }
        private bool IsWeekEnd(DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        }
        private bool IsWorkingDay(DateTime date)
        {
            return WorkingDays.Contains(date.Date) || !IsWeekEnd(date) && !NotWorkingDays.Contains(date.Date);
        }
        private DateTime GetPreviousAndWorkingDate()
        {
            DateTime prevDate = DateTime.Now.Subtract(new TimeSpan(1, 0, 0, 0)).Date;

            while (!IsWorkingDay(prevDate))
            {
                prevDate = prevDate.Subtract(new TimeSpan(1)).Date;
            }

            return prevDate;
        }
        #endregion
        /// <summary>
        /// Processing daily report
        /// </summary>
        public async Task ProcessAggregatedDailyFile()
        {
            if(IsWorkingDay(DateTime.Now))
            {
                // Load excel template1
                Console.WriteLine("Load -> " + Template1Path);
                Excel.Application template1App = new Excel.Application();
                Excel.Workbook template1Wb = template1App.Workbooks.Open(Template1Path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet template1Sheet_GFU = (Excel.Worksheet)template1Wb.Sheets["ГФУ"];
                Excel.Range range1_GFU = template1Sheet_GFU.Range["C6", "F29"];
                Excel.Worksheet template1Sheet_GU = (Excel.Worksheet)template1Wb.Sheets["ГУ"];
                Excel.Range range1_GU = template1Sheet_GU.Range["C6", "F29"];

                // Load excel template2
                Console.WriteLine("Load -> " + Template2Path);
                Excel.Application template2App = new Excel.Application();
                Excel.Workbook template2Wb = 
                    template2App.Workbooks.Open(Template2Path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet template2Sheet = (Excel.Worksheet)template2Wb.Sheets["Main"];
                Excel.Range range2 = template2Sheet.Range["C3", "E25"];

                if (File.Exists(AggregateDailyArchiveFileName))
                    File.Delete(AggregateDailyArchiveFileName);

                if (File.Exists(AggregateDailyArchiveLogFileName))
                    File.Delete(AggregateDailyArchiveLogFileName);

                await Task.Run(
                    () => Aggregate(Template1Path, Template2Path, range1_GFU, range1_GU, range2))
                    .ContinueWith(a => template1Wb.SaveAs(AggregateDailyArchiveFileName))
                    .ContinueWith(b => template2Wb.SaveAs(AggregateDailyArchiveLogFileName))
                    .ContinueWith(c => Dispose(template1App, template1Wb, new Tuple<Excel.Worksheet, Excel.Worksheet>(template1Sheet_GFU, template1Sheet_GU)))
                    .ContinueWith(d => Dispose(template2App, template2Wb, new Tuple<Excel.Worksheet, Excel.Worksheet>(template2Sheet, template2Sheet)))
                    //.ContinueWith(e => new OutlookSender(AggregateDailyArchiveFileName, "Отчёт за " + Convert.ToDateTime(GetPreviousAndWorkingDate()).ToString("dd.MM.yyyy"), "Журнал инцидентов", EmailSubject, Email, logger).SendEmail())
                    .ContinueWith(e => new MailSender(logger).SendEmail(AggregateDailyArchiveFileName, "Отчёт за " + Convert.ToDateTime(GetPreviousAndWorkingDate()).ToString("dd.MM.yyyy"), EmailSubject, new List<string>() { Email1, Email2 }))
                    .ContinueWith(f => Cleanup());

                logger.LogMessage(string.Empty, true);
                File.Copy(MainLoggerPath, AdditionalLoggerPath, true);

                //Console.ReadLine();
            }
            else
            {
                logger.LogMessage(" -> Not working day!");
            }
        }
        #region Проверки
        private Tuple<bool, bool> FileExistAndEmpty(string path)
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
        private string FixedIncidentType(string str)
        {
            string fixedStr = string.Empty;

            if(str.Contains("Нет") && str.Contains("инц"))
            {
                logger.LogMessage("Fixing incident type -> " + str);
                fixedStr = "Нет инцидентов";
            }

            return fixedStr;
        }
        private bool DateExist(string date, string region)
        {
            Console.WriteLine("[DateExist] COMPARE: for {0} - {1} AND {2}", region, date, GetPreviousAndWorkingDate().ToShortDateString());
            logger.LogMessage("[DateExist] COMPARE: for " + region + " - " + date + " AND " + GetPreviousAndWorkingDate().ToShortDateString());

            /*
            try
            {
                parsedDate = DateTime.Parse(date);
            }
            catch (ArgumentNullException ex)
            {
                Console.WriteLine("!!! ArgumentNullException !!!" + ex.Message + " || " + date);
                return false;
            }
            catch(FormatException ex)
            {
                Console.WriteLine("!!! FormatException !!!" + ex.Message + " || " + date);
                return false;
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("!!! ArgumentException !!!" + ex.Message + " || " + date);
                return false;
            }
            */
            if (date == "Null" || date == "Empty" || date == "Skip")
            {
                return false;
            }
            else if (DateTime.TryParse(date, out DateTime res) && GetPreviousAndWorkingDate() == res)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private Color SetCellColor(bool state, int counter, bool skip)
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
        private string SetLogString(int counter, string msg, bool skip)
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
        private bool IsStringsIsEmptyOrNull(string str1, string str2)
        {
            bool result = (str1 == "Empty" || str2 == "Empty") || (string.IsNullOrEmpty(str1) || string.IsNullOrEmpty(str2)) ? true : false;
            //logger.LogMessage("[IsStringsIsEmptyOrNull] " + result + " str1: " + str1 + " str2: " + str2);
            Console.WriteLine("STR1 {0}, STR2 {1}", str1, str2);
            return result;
        }
        #endregion
        private void Aggregate(string path1, string path2, Excel.Range rangeGFU, Excel.Range rangeGU, Excel.Range rangeLog)
        {
            int counter = 0;
            int regionPathCollectionCount;

            string outputStr1 = string.Empty;
            string outputStr2 = string.Empty;

            bool gfuAndGuFileIsExist;
            bool guCityFileExist;

            bool gfuDateIsCorrect;
            bool guDateIsCorrect;
            bool guCityDateIsCorrect;

            bool gfuStringsIsEmpty;
            bool guStringsIsEmpty;
            bool guCityStringsIsEmpty;

            Tuple<List<string>, List<string>> excelData_FULL;

            List<string> excelData_GFU = new List<string>();
            List<string> excelData_GU = new List<string>();

            foreach (ConnectionManagerPathElement pathElement in RegionPathCollection.Cast<ConnectionManagerPathElement>().Take(RegionPathCollection.Count - 1))
            {
                excelData_FULL = LoadExcel(pathElement);

                var fileIsOk = FileExistAndEmpty(pathElement.Path);

                gfuAndGuFileIsExist = fileIsOk.Item1 && !fileIsOk.Item2;

                logger.LogMessage(pathElement + " -> " + gfuAndGuFileIsExist);

                gfuDateIsCorrect = DateExist(excelData_FULL.Item1.ElementAt(3), counter.ToString());
                guDateIsCorrect = DateExist(excelData_FULL.Item2.ElementAt(3), counter.ToString());
                gfuStringsIsEmpty = IsStringsIsEmptyOrNull(excelData_FULL.Item1.ElementAt(4), excelData_FULL.Item1.ElementAt(5));
                guStringsIsEmpty = IsStringsIsEmptyOrNull(excelData_FULL.Item2.ElementAt(4), excelData_FULL.Item2.ElementAt(5));

                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 1])
                    .Interior.Color = ColorTranslator.ToOle(SetCellColor(gfuAndGuFileIsExist, counter, false));

                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 2])
                    .Interior.Color = ColorTranslator.ToOle(SetCellColor(gfuAndGuFileIsExist && gfuDateIsCorrect, counter, false));
                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 3])
                    .Interior.Color = ColorTranslator.ToOle(SetCellColor(gfuAndGuFileIsExist && guDateIsCorrect, counter, true));

                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 4])
                    .Interior.Color = ColorTranslator.ToOle(SetCellColor(gfuAndGuFileIsExist && !gfuStringsIsEmpty, counter, false));
                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 5])
                    .Interior.Color = ColorTranslator.ToOle(SetCellColor(gfuAndGuFileIsExist && !guStringsIsEmpty, counter, true));

                if (gfuAndGuFileIsExist)
                {
                    ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 1]).Value2 = SetData(excelData_FULL.Item1.ElementAt(2));
                    ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 1]).Value2 = SetData(excelData_FULL.Item2.ElementAt(2));

                    ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = SetLogString(counter, excelData_FULL.Item1.ElementAt(2), false);
                    ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = SetLogString(counter, excelData_FULL.Item2.ElementAt(2), false);

                    // Дата, время
                    ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 2]).Value2 = GetPreviousAndWorkingDate().ToShortDateString();
                    ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 2]).Value2 = GetPreviousAndWorkingDate().ToShortDateString();

                    // Описание обстоятельств инцидента
                    ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 3]).Value2 = SetData(excelData_FULL.Item1.ElementAt(4));
                    ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 3]).Value2 = SetData(excelData_FULL.Item2.ElementAt(4));

                    // Какие меры приняты
                    ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 4]).Value2 = SetData(excelData_FULL.Item1.ElementAt(5));
                    ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 4]).Value2 = SetData(excelData_FULL.Item2.ElementAt(5));

                    if (counter == 0 || counter == 2 || counter == 22)
                    {
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 1]).Value2 = "-";
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 2]).Value2 = "-";
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 3]).Value2 = "-";
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 4]).Value2 = "-";
                    }
                    else
                    {
                    }
                }
                // Файла нет.
                else
                {
                    if (counter == 0 || counter == 2 || counter == 22)
                    {
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 1]).Value2 = "-";
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 2]).Value2 = "-";
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 3]).Value2 = "-";
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 4]).Value2 = "-";

                        //((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 1]).Value2 = "Нет инцидентов";
                        ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 2]).Value2 = GetPreviousAndWorkingDate().ToShortDateString();

                        if(fileIsOk.Item2)
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = GetEnumDescription(ErrorMsgs.FileIsEmpty);
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = GetEnumDescription(ErrorMsgs.FileIsEmpty);
                        }
                        else
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = GetEnumDescription(ErrorMsgs.MissingFile);
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = GetEnumDescription(ErrorMsgs.MissingFile);
                        }
                        

                        //((Excel.Range)rangeLog.Rows.Cells[counter + 1, 9]).Value2 = Path.GetDirectoryName(pathElement.Path);
                    }
                    else if(counter == 23)
                    {
                        ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 1]).Value2 = "Нет инцидентов";
                        ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 2]).Value2 = GetPreviousAndWorkingDate().ToShortDateString();

                        if (fileIsOk.Item2)
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = GetEnumDescription(ErrorMsgs.FileIsEmpty);
                        }
                        else
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = GetEnumDescription(ErrorMsgs.MissingFile);
                        }
                    }
                    else
                    {
                        //((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 1]).Value2 = "Нет инцидентов";
                        ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 2]).Value2 = GetPreviousAndWorkingDate().ToShortDateString();
                        //((Excel.Range)rangeGU.Rows.Cells[counter + 1, 1]).Value2 = "Нет инцидентов";
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 2]).Value2 = GetPreviousAndWorkingDate().ToShortDateString();

                        if (fileIsOk.Item2)
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = GetEnumDescription(ErrorMsgs.FileIsEmpty);
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = GetEnumDescription(ErrorMsgs.FileIsEmpty);
                        }
                        else
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = GetEnumDescription(ErrorMsgs.MissingFile);
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = GetEnumDescription(ErrorMsgs.MissingFile);
                        }

                        //((Excel.Range)rangeLog.Rows.Cells[counter + 1, 9]).Value2 = Path.GetDirectoryName(pathElement.Path);
                    }
                }

                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 8]).Value2 = GetFolderModifiedDate(Path.GetDirectoryName(pathElement.Path));

                logger.LogMessage(excelData_FULL.Item1.ElementAt(1) + " -> Done.");

                counter++;
            }

            // Обработка ГУ вне цикла.
            ConnectionManagerPathElement mogGuPath = RegionPathCollection.Cast<ConnectionManagerPathElement>().Last();
            List<string> mogGuData = LoadExcel(mogGuPath).Item2;

            var isGuFileOk = FileExistAndEmpty(mogGuPath.Path);

            guCityFileExist = isGuFileOk.Item1 && !isGuFileOk.Item2;
            guCityDateIsCorrect = DateExist(mogGuData.ElementAt(3), counter.ToString());
            guCityStringsIsEmpty = IsStringsIsEmptyOrNull(mogGuData.ElementAt(4), mogGuData.ElementAt(5));

            regionPathCollectionCount = RegionPathCollection.Cast<ConnectionManagerPathElement>().Count();

            // Проверки для лога.
            ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 1]).Interior.Color = ColorTranslator.ToOle(SetCellColor(guCityFileExist, counter, false));
            ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 3]).Interior.Color = ColorTranslator.ToOle(SetCellColor(guCityFileExist && guCityDateIsCorrect, counter, false));
            ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 5]).Interior.Color = ColorTranslator.ToOle(SetCellColor(guCityFileExist && guCityDateIsCorrect && !guCityStringsIsEmpty, counter, false));

            ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 8]).Value2 = GetFolderModifiedDate(Path.GetDirectoryName(mogGuPath.Path));

            if (guCityFileExist)
            {
                if (guCityDateIsCorrect)
                {
                    // Инцидент (если не Empty, то ставим Технический).
                    if (mogGuData.ElementAt(2) != "Empty" && guCityStringsIsEmpty == false)
                    {
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 1]).Value2 = mogGuData.ElementAt(2);
                        ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 7]).Value2 = mogGuData.ElementAt(2);

                        // Дата.
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 2]).Value2 = mogGuData.ElementAt(3);

                        // Обстоятельства.
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 3]).Value2 = mogGuData.ElementAt(4);

                        // Меры.
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 4]).Value2 = mogGuData.ElementAt(5);
                    }
                    // Если поля обстоятельства и меры не Empty.
                    else if (mogGuData.ElementAt(2) == "Empty" && guCityStringsIsEmpty == false)
                    {
                        ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 7]).Value2 = SetLogString(counter, GetEnumDescription(ErrorMsgs.MissingIncident), true);
                        //((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 7]).Value2 = "Технический (wrong type)";
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 1]).Value2 = "Технический";

                        // Дата.
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 2]).Value2 = mogGuData.ElementAt(3);

                        // Обстоятельства.
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 3]).Value2 = mogGuData.ElementAt(4);

                        // Меры.
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 4]).Value2 = mogGuData.ElementAt(5);
                    }
                    else
                    {
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 1]).Value2 = "Нет инцидентов";
                        ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 7]).Value2 = GetEnumDescription(ErrorMsgs.MissingStrings);

                        // Дата.
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 2]).Value2 = GetPreviousAndWorkingDate().ToShortDateString();
                    }

                    
                }
                else
                {
                    ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = SetLogString(regionPathCollectionCount, GetEnumDescription(ErrorMsgs.WrongDate), true);
                }
            }
            else
            {
                ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 1]).Value2 = "Нет инцидентов";

                if (isGuFileOk.Item2)
                {
                    ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = SetLogString(regionPathCollectionCount, GetEnumDescription(ErrorMsgs.FileIsEmpty), true);
                }
                else
                {
                    ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = SetLogString(regionPathCollectionCount, GetEnumDescription(ErrorMsgs.MissingFile), true);
                }

                // Дата.
                ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 2]).Value2 = GetPreviousAndWorkingDate().ToShortDateString();
                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 8]).Value2 = GetFolderModifiedDate(Path.GetDirectoryName(mogGuPath.Path));

                //((Excel.Range)rangeLog.Rows.Cells[counter + 1, 9]).Value2 = Path.GetDirectoryName(mogGuPath.Path);
            }

            logger.LogMessage(mogGuData.ElementAt(1) + " -> Done.");

            Directory.CreateDirectory(Path.GetDirectoryName(AggregateDailyArchiveFileName));
            Directory.CreateDirectory(Path.GetDirectoryName(AggregateDailyArchiveLogFileName));

            logger.LogMessage("Aggregate" + " -> Done.");
        }
        private string SetData(string rawStr)
        {
            if(rawStr.Equals("Empty") || string.IsNullOrEmpty(rawStr) || string.IsNullOrWhiteSpace(rawStr))
            {
                return string.Empty;
            }
            else
            {
                return rawStr;
            }
        }
        private void Cleanup()
        {
            try
            {
                foreach (ConnectionManagerPathElement pathElement in RegionPathCollection)
                {
                    if (!Directory.Exists(RawArchivePath + @"\" + pathElement.Name))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(RawArchivePath + @"\" + pathElement.Name + @"\" + "Журнал инцидентов.xlsx"));
                        MoveFile(pathElement);
                    }
                    else
                    {
                        MoveFile(pathElement);
                    }
                }

                Console.WriteLine("Cleanup complete");
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex);
            }
        }
        private void MoveFile(ConnectionManagerPathElement pathElement)
        {
            if (File.Exists(pathElement.Path))
            {
                logger.LogMessage("[MoveFile] " + pathElement.Path + " -> " + RawArchivePath + @"\" + pathElement.Name);
                //logger.LogMessage("File exist " + pathElement.Path + ". Moving file -> " + RawArchivePath + @"\" + pathElement.Name);
                File.Move(pathElement.Path, RawArchivePath + @"\" + pathElement.Name + @"\" + "Журнал инцидентов.xlsx");
            }
            else
            {
                //logger.LogMessageToFile("File missing " + pathElement.Path + ". Skipping file moving.");
            }
        }
        private Tuple<List<string>, List<string>> LoadExcel(ConnectionManagerPathElement pathElement)
        {
            List<string> data_GFU = new List<string>();
            List<string> data_GU = new List<string>();

            Excel.Application rApp;
            Excel.Workbook rWb;
            Excel.Worksheet rSheet_GFU;
            Excel.Worksheet rSheet_GU;
            Excel.Range range_GFU;
            Excel.Range range_GU;

            var isFileOk = FileExistAndEmpty(pathElement.Path);

            if (isFileOk.Item1 && !isFileOk.Item2)
            {
                rApp = new Excel.Application();
                rWb = rApp.Workbooks.Open(pathElement.Path);

                rSheet_GFU = (Excel.Worksheet)rWb.Sheets["ГФУ"];
                range_GFU = rSheet_GFU.Range["A6", "F6"];

                rSheet_GU = (Excel.Worksheet)rWb.Sheets["ГУ"];
                range_GU = rSheet_GU.Range["A6", "F6"];

                //Console.WriteLine("[LoadExcel] " + pathElement.Path);
                logger.LogMessage("[LoadExcel] " + pathElement.Path);

                for (int i = 0; i < 6; i++)
                {
                    if (Convert.ToString(((Excel.Range)range_GFU.Rows.Cells[1, i + 1]).Value2) == null)
                    {
                        data_GFU.Add("Empty");
                    }
                    else
                    {
                        // Date.
                        if (i == 3)
                        {
                            string valGFU = Convert.ToString(((Excel.Range)range_GFU.Rows.Cells[1, i + 1]).Value2);
                            valGFU = valGFU.Trim();

                            //Console.WriteLine("DT parse " + valGFU + " and indexed val: " + valGFU[2]);

                            if (valGFU[2] == '.')
                            {
                                Console.WriteLine("DEFAULT DATE");

                                string raw = valGFU.Substring(0, 10);

                                Console.WriteLine("RAW SUB " + raw);

                                DateTime parsedDt = DateTime.Parse(raw);
                                parsedDt = parsedDt.Date;

                                data_GFU.Add(parsedDt.ToShortDateString());

                                Console.WriteLine("DEFAULT DATE PARSED: " + parsedDt.ToShortDateString());
                            }
                            /*
                            else if (valGFU[2] == '/')
                            {
                                //data_GFU.Add(DateParser(valGFU).ToShortDateString());
                            }
                            */
                            else
                            {
                                Console.WriteLine("OL DATE");

                                double d = double.MinValue;
 
                                DateTime conv;

                                try
                                {
                                    bool isOlDateParsed = double.TryParse(valGFU, out d);

                                    conv = DateTime.FromOADate(d);

                                    if (isOlDateParsed)
                                    {
                                        logger.LogMessage("[LoadExcel][GFU] OL DATE PARSED: " + conv.ToShortDateString());
                                        data_GFU.Add(conv.ToShortDateString());
                                    }
                                    else
                                    {
                                        logger.LogMessage("[LoadExcel][GFU] OL DATE PARSER ERROR: " + valGFU);
                                        data_GFU.Add("Empty");
                                    }
                                }
                                catch (ArgumentException ex)
                                {
                                    Console.WriteLine("!!! Argument exception !!! (GU) " + ex.Message + " " + d);
                                    logger.LogMessage("[LoadExcel][GFU] OL DATE Parse exception." + valGFU);
                                    data_GFU.Add("Empty");
                                }

                            }
                        }
                        else
                        {
                            data_GFU.Add(Convert.ToString(((Excel.Range)range_GFU.Rows.Cells[1, i + 1]).Value2));
                        }
                    }
                    if (Convert.ToString(((Excel.Range)range_GU.Rows.Cells[1, i + 1]).Value2) == null)
                    {
                        data_GU.Add("Empty");
                    }
                    else
                    {
                        // Date.
                        if (i == 3)
                        {
                            if (Convert.ToString(((Excel.Range)range_GU.Rows.Cells[1, i + 1]).Value2) != "-")
                            {
                                string valGU = Convert.ToString((((Excel.Range)range_GU.Rows.Cells[1, i + 1])).Value2);
                                valGU = valGU.Trim();

                                
                                //Console.WriteLine("DT parse " + valGU + " and indexed val: " + valGU[2]);

                                if (valGU[2] == '.')
                                {
                                    Console.WriteLine("DEFAULT DATE");

                                    string raw = valGU.Substring(0, 10);

                                    Console.WriteLine("RAW SUB " + raw);

                                    DateTime parsedDt = DateTime.Parse(raw);
                                    parsedDt = parsedDt.Date;

                                    data_GU.Add(parsedDt.ToShortDateString());

                                    logger.LogMessage("DEFAULT DATE PARSED: " + parsedDt.ToShortDateString());
                                }
                                /*
                                else if (valGU[2] == '/')
                                {
                                    //data_GU.Add(DateParser(valGU).ToShortDateString());
                                }
                                */
                                else
                                {
                                    Console.WriteLine("OL DATE");

                                    double d = double.MinValue;

                                    DateTime conv;

                                    try
                                    {
                                        bool isOlDateParsed = double.TryParse(valGU, out d);

                                        conv = DateTime.FromOADate(d);

                                        if (isOlDateParsed)
                                        {
                                            logger.LogMessage("[LoadExcel][GU] OL DATE PARSED: " + conv.ToShortDateString());
                                            data_GU.Add(conv.ToShortDateString());
                                        }
                                        else
                                        {
                                            logger.LogMessage("[LoadExcel][GU] OL DATE PARSER ERROR: " + valGU);
                                            data_GU.Add("Empty");
                                        }
                                    }
                                    catch (ArgumentException ex)
                                    {
                                        Console.WriteLine("!!! Argument exception !!! (GU) " + ex.Message + " " + d);
                                        logger.LogMessage("[LoadExcel][GU] OL DATE Parse exception." + valGU);
                                        data_GU.Add("Empty");
                                    }
                                }
                            }
                            else
                            {
                                data_GU.Add("Skip");
                            }
                        }
                        else
                        {
                            data_GU.Add(Convert.ToString(((Excel.Range)range_GU.Rows.Cells[1, i + 1]).Value2));
                        }
                    }
                }

                Dispose(rApp, rWb, new Tuple<Excel.Worksheet, Excel.Worksheet>(rSheet_GFU, rSheet_GU));
            }
            else
            {
                for (int i = 0; i < 6; i++)
                {
                    data_GFU.Add("Null");
                    data_GU.Add("Null");
                }
            }

            return new Tuple<List<string>, List<string>>(data_GFU, data_GU);
        }
        private DateTime DateParser(string rawDate)
        {
            DateTime dt = GetPreviousAndWorkingDate();
            string parsedValue = rawDate.Substring(0, rawDate.IndexOf(','));

            if (DateTime.TryParse(parsedValue, out dt))
            {
                logger.LogMessage("[DATE PARSER] Ol date parsed: " + dt.ToShortDateString() + ".");
            }
            else
            {
                logger.LogMessage("[DATE PARSER] Ol date parser error for " + parsedValue + ". Using the previous working date.");
            }

            return dt;
        }
        private string GetEnumDescription(Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }
        private void Dispose(Excel.Application iApp, Excel.Workbook iWb, Tuple<Excel.Worksheet, Excel.Worksheet> nSheets)
        {
            // Cleanup
            if (nSheets != null)
            {
                if (nSheets.Item1 != null)
                    Marshal.ReleaseComObject(nSheets.Item1);

                if (nSheets.Item2 != null)
                    Marshal.ReleaseComObject(nSheets.Item1);
            }
            if (iWb != null)
            {
                iWb.Close(true);
                Marshal.ReleaseComObject(iWb);
            }
            if (iApp != null)
            {
                iApp.Quit();
                Marshal.ReleaseComObject(iApp);
            }
            iApp = null;
            iWb = null;
            nSheets = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        private string GetFolderModifiedDate(string folder)
        {
            string result = string.Empty;

            try
            {
                if (!Directory.Exists(folder))
                {
                    logger.LogMessage("[GetFolderModifiedDate] Directory -> " + folder + " didn't exist.");
                    result = "-";
                }
                else
                {
                    // Get the creation time of a well-known directory.
                    result = Directory.GetLastWriteTime(folder).ToString();
                    logger.LogMessage("The last write time for " + folder + " directory was " + result + ".");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
                result = "Error";
            }

            return result;
        }
    }
}