using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using Incidents.Interfaces;

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

        private ILogger logger;
        private readonly DateManager dateManager;
        private readonly IPathManager pathManager;
        private readonly ICleanupManager cleanupManager;
        private readonly IStringHelper stringHelper;

        public DailyReport()
        {
            dateManager = new DateManager();
            pathManager = new PathManager(dateManager);
            cleanupManager = new CleanupHelper(logger, pathManager);
            stringHelper = new StringHelper(logger, dateManager);

            // Load dates
            dateManager.LoadDates();

            Console.WriteLine("Load logger");
            // Load logger
            logger = new LogToTxt(new List<string>() { pathManager.MainLoggerPath });
            Console.WriteLine("Load complete");
        }
        
        /// <summary>
        /// Processing daily report
        /// </summary>
        public async Task ProcessAggregatedDailyFile()
        {
            if(dateManager.IsWorkingDay(DateTime.Now))
            {
                // Load excel template1
                Console.WriteLine("Load -> " + pathManager.Template1Path);
                Excel.Application template1App = new Excel.Application();
                Excel.Workbook template1Wb = template1App.Workbooks
                    .Open(pathManager.Template1Path, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing);
                Excel.Worksheet template1Sheet_GFU = (Excel.Worksheet)template1Wb.Sheets["ГФУ"];
                Excel.Range range1_GFU = template1Sheet_GFU.Range["C6", "F29"];
                Excel.Worksheet template1Sheet_GU = (Excel.Worksheet)template1Wb.Sheets["ГУ"];
                Excel.Range range1_GU = template1Sheet_GU.Range["C6", "F29"];

                // Load excel template2
                Console.WriteLine("Load -> " + pathManager.Template2Path);
                Excel.Application template2App = new Excel.Application();
                Excel.Workbook template2Wb = 
                    template2App.Workbooks.Open(pathManager.Template2Path, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing);
                Excel.Worksheet template2Sheet = (Excel.Worksheet)template2Wb.Sheets["Main"];
                Excel.Range range2 = template2Sheet.Range["C3", "E25"];

                if (File.Exists(pathManager.AggregateDailyArchiveFileName))
                    File.Delete(pathManager.AggregateDailyArchiveFileName);

                if (File.Exists(pathManager.AggregateDailyArchiveLogFileName))
                    File.Delete(pathManager.AggregateDailyArchiveLogFileName);

                await Task.Run(
                    () => Aggregate(pathManager.Template1Path, pathManager.Template2Path, range1_GFU, range1_GU, range2))
                    .ContinueWith(a => template1Wb.SaveAs(pathManager.AggregateDailyArchiveFileName))
                    .ContinueWith(b => template2Wb.SaveAs(pathManager.AggregateDailyArchiveLogFileName))
                    .ContinueWith(c => cleanupManager.Dispose(template1App, template1Wb, 
                    new Tuple<Excel.Worksheet, Excel.Worksheet>(template1Sheet_GFU, template1Sheet_GU)))
                    .ContinueWith(d => cleanupManager.Dispose(template2App, template2Wb, 
                    new Tuple<Excel.Worksheet, Excel.Worksheet>(template2Sheet, template2Sheet)))
                    .ContinueWith(e => 
                    new MailSender(logger).SendEmail(pathManager.AggregateDailyArchiveFileName, "Отчёт за " 
                    + Convert.ToDateTime(dateManager.GetPreviousAndWorkingDate())
                    .ToString("dd.MM.yyyy"), pathManager.EmailSubject, new List<string>() { pathManager.Email1, pathManager.Email2 }))
                    .ContinueWith(f => cleanupManager.Cleanup());

                logger.LogMessage(string.Empty, true);
                File.Copy(pathManager.MainLoggerPath, pathManager.AdditionalLoggerPath, true);

                //Console.ReadLine();
            }
            else
            {
                logger.LogMessage(" -> Not working day!");
            }
        }
        // TODO. Сделано на скорую руку. Найти лучший способ прохода по Excel-файлам.
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

            foreach (ConnectionManagerPathElement pathElement in 
                pathManager.RegionPathCollection.Cast<ConnectionManagerPathElement>().Take(pathManager.RegionPathCollection.Count - 1))
            {
                excelData_FULL = LoadExcel(pathElement);

                var fileIsOk = stringHelper.FileExistAndEmpty(pathElement.Path);

                gfuAndGuFileIsExist = fileIsOk.Item1 && !fileIsOk.Item2;

                logger.LogMessage(pathElement + " -> " + gfuAndGuFileIsExist);

                gfuDateIsCorrect = stringHelper.DateExist(excelData_FULL.Item1.ElementAt(3), counter.ToString());
                guDateIsCorrect = stringHelper.DateExist(excelData_FULL.Item2.ElementAt(3), counter.ToString());
                gfuStringsIsEmpty = stringHelper.IsStringsIsEmptyOrNull(excelData_FULL.Item1.ElementAt(4), excelData_FULL.Item1.ElementAt(5));
                guStringsIsEmpty = stringHelper.IsStringsIsEmptyOrNull(excelData_FULL.Item2.ElementAt(4), excelData_FULL.Item2.ElementAt(5));

                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 1])
                    .Interior.Color = ColorTranslator.ToOle(stringHelper.SetCellColor(gfuAndGuFileIsExist, counter, false));

                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 2])
                    .Interior.Color = ColorTranslator.ToOle(stringHelper.SetCellColor(gfuAndGuFileIsExist && gfuDateIsCorrect, counter, false));
                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 3])
                    .Interior.Color = ColorTranslator.ToOle(stringHelper.SetCellColor(gfuAndGuFileIsExist && guDateIsCorrect, counter, true));

                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 4])
                    .Interior.Color = ColorTranslator.ToOle(stringHelper.SetCellColor(gfuAndGuFileIsExist && !gfuStringsIsEmpty, counter, false));
                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 5])
                    .Interior.Color = ColorTranslator.ToOle(stringHelper.SetCellColor(gfuAndGuFileIsExist && !guStringsIsEmpty, counter, true));

                if (gfuAndGuFileIsExist)
                {
                    ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 1]).Value2 = dateManager.SetData(excelData_FULL.Item1.ElementAt(2));
                    ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 1]).Value2 = dateManager.SetData(excelData_FULL.Item2.ElementAt(2));

                    ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = 
                        stringHelper.SetLogString(counter, excelData_FULL.Item1.ElementAt(2), false);
                    ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = 
                        stringHelper.SetLogString(counter, excelData_FULL.Item2.ElementAt(2), false);

                    // Дата, время
                    ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 2]).Value2 = dateManager.GetPreviousAndWorkingDate().ToShortDateString();
                    ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 2]).Value2 = dateManager.GetPreviousAndWorkingDate().ToShortDateString();

                    // Описание обстоятельств инцидента
                    ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 3]).Value2 = dateManager.SetData(excelData_FULL.Item1.ElementAt(4));
                    ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 3]).Value2 = dateManager.SetData(excelData_FULL.Item2.ElementAt(4));

                    // Какие меры приняты
                    ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 4]).Value2 = dateManager.SetData(excelData_FULL.Item1.ElementAt(5));
                    ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 4]).Value2 = dateManager.SetData(excelData_FULL.Item2.ElementAt(5));

                    if (counter == 0 || counter == 2 || counter == 22)
                    {
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 1]).Value2 = "-";
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 2]).Value2 = "-";
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 3]).Value2 = "-";
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 4]).Value2 = "-";
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
                        ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 2])
                            .Value2 = dateManager.GetPreviousAndWorkingDate().ToShortDateString();

                        if(fileIsOk.Item2)
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = ErrorMsgs.FileIsEmpty.GetEnumDescription();
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = ErrorMsgs.FileIsEmpty.GetEnumDescription();
                        }
                        else
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = ErrorMsgs.MissingFile.GetEnumDescription();
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = ErrorMsgs.MissingFile.GetEnumDescription();
                        }
                    }
                    else if(counter == 23)
                    {
                        ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 1]).Value2 = "Нет инцидентов";
                        ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 2]).Value2 = dateManager.GetPreviousAndWorkingDate().ToShortDateString();

                        if (fileIsOk.Item2)
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = ErrorMsgs.FileIsEmpty.GetEnumDescription();
                        }
                        else
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = ErrorMsgs.MissingFile.GetEnumDescription();
                        }
                    }
                    else
                    {
                        ((Excel.Range)rangeGFU.Rows.Cells[counter + 1, 2]).Value2 = 
                            dateManager.GetPreviousAndWorkingDate().ToShortDateString();
                        ((Excel.Range)rangeGU.Rows.Cells[counter + 1, 2]).Value2 = 
                            dateManager.GetPreviousAndWorkingDate().ToShortDateString();

                        if (fileIsOk.Item2)
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = ErrorMsgs.FileIsEmpty.GetEnumDescription();
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = ErrorMsgs.FileIsEmpty.GetEnumDescription();
                        }
                        else
                        {
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 6]).Value2 = ErrorMsgs.MissingFile.GetEnumDescription();
                            ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = ErrorMsgs.MissingFile.GetEnumDescription();
                        }
                    }
                }

                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 8]).Value2 = 
                    dateManager.GetFolderModifiedDate(Path.GetDirectoryName(pathElement.Path));

                logger.LogMessage(excelData_FULL.Item1.ElementAt(1) + " -> Done.");

                counter++;
            }

            // Обработка ГУ вне цикла.
            ConnectionManagerPathElement mogGuPath = pathManager.RegionPathCollection.Cast<ConnectionManagerPathElement>().Last();
            List<string> mogGuData = LoadExcel(mogGuPath).Item2;

            var isGuFileOk = stringHelper.FileExistAndEmpty(mogGuPath.Path);

            guCityFileExist = isGuFileOk.Item1 && !isGuFileOk.Item2;
            guCityDateIsCorrect = stringHelper.DateExist(mogGuData.ElementAt(3), counter.ToString());
            guCityStringsIsEmpty = stringHelper.IsStringsIsEmptyOrNull(mogGuData.ElementAt(4), mogGuData.ElementAt(5));

            regionPathCollectionCount = pathManager.RegionPathCollection.Cast<ConnectionManagerPathElement>().Count();

            // Проверки для лога.
            ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 1]).Interior.Color 
                = ColorTranslator.ToOle(stringHelper.SetCellColor(guCityFileExist, counter, false));
            ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 3]).Interior.Color 
                = ColorTranslator.ToOle(stringHelper.SetCellColor(guCityFileExist && guCityDateIsCorrect, counter, false));
            ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 5]).Interior.Color 
                = ColorTranslator.ToOle(stringHelper.SetCellColor(guCityFileExist && guCityDateIsCorrect && !guCityStringsIsEmpty, counter, false));

            ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 8]).Value2 = 
                dateManager.GetFolderModifiedDate(Path.GetDirectoryName(mogGuPath.Path));

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
                        ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 7]).Value2 
                            = stringHelper.SetLogString(counter, ErrorMsgs.MissingIncident.GetEnumDescription(), true);
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
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 1])
                            .Value2 = "Нет инцидентов";
                        ((Excel.Range)rangeLog.Rows.Cells[regionPathCollectionCount, 7])
                            .Value2 = ErrorMsgs.MissingStrings.GetEnumDescription();

                        // Дата.
                        ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 2])
                            .Value2 = dateManager.GetPreviousAndWorkingDate().ToShortDateString();
                    }

                    
                }
                else
                {
                    ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = 
                        stringHelper.SetLogString(regionPathCollectionCount, ErrorMsgs.WrongDate.GetEnumDescription(), true);
                }
            }
            else
            {
                ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 1]).Value2 = "Нет инцидентов";

                if (isGuFileOk.Item2)
                {
                    ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = 
                        stringHelper.SetLogString(regionPathCollectionCount, ErrorMsgs.FileIsEmpty.GetEnumDescription(), true);
                }
                else
                {
                    ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 7]).Value2 = 
                        stringHelper.SetLogString(regionPathCollectionCount, ErrorMsgs.MissingFile.GetEnumDescription(), true);
                }

                // Дата.
                ((Excel.Range)rangeGU.Rows.Cells[regionPathCollectionCount - 1, 2]).Value2 = dateManager
                    .GetPreviousAndWorkingDate().ToShortDateString();
                ((Excel.Range)rangeLog.Rows.Cells[counter + 1, 8]).Value2 = 
                    dateManager.GetFolderModifiedDate(Path.GetDirectoryName(mogGuPath.Path));
            }

            logger.LogMessage(mogGuData.ElementAt(1) + " -> Done.");

            Directory.CreateDirectory(Path.GetDirectoryName(pathManager.AggregateDailyArchiveFileName));
            Directory.CreateDirectory(Path.GetDirectoryName(pathManager.AggregateDailyArchiveLogFileName));

            logger.LogMessage("Aggregate" + " -> Done.");
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

            var isFileOk = stringHelper.FileExistAndEmpty(pathElement.Path);

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

                cleanupManager.Dispose(rApp, rWb, new Tuple<Excel.Worksheet, Excel.Worksheet>(rSheet_GFU, rSheet_GU));
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
    }
}