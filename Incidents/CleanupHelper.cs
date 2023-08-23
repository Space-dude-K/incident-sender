using Incidents.Interfaces;
using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Incidents
{
    public class CleanupHelper : ICleanupManager
    {
        private readonly ILogger logger;
        private readonly IPathManager pathManager;

        public CleanupHelper(ILogger logger, IPathManager pathManager)
        {
            this.logger = logger;
            this.pathManager = pathManager;
        }

        public void MoveFile(ConnectionManagerPathElement pathElement)
        {
            if (File.Exists(pathElement.Path))
            {
                logger.LogMessage("[MoveFile] " + pathElement.Path + " -> " + pathManager.RawArchivePath + @"\" + pathElement.Name);
                //logger.LogMessage("File exist " + pathElement.Path + ". Moving file -> " + pathManager.RawArchivePath + @"\" + pathElement.Name);
                File.Move(pathElement.Path, pathManager.RawArchivePath + @"\" + pathElement.Name + @"\" + "Журнал инцидентов.xlsx");
            }
            else
            {
                //logger.LogMessageToFile("File missing " + pathElement.Path + ". Skipping file moving.");
            }
        }
        public void Cleanup()
        {
            try
            {
                foreach (ConnectionManagerPathElement pathElement in pathManager.RegionPathCollection)
                {
                    if (!Directory.Exists(pathManager.RawArchivePath + @"\" + pathElement.Name))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(pathManager.RawArchivePath
                            + @"\" + pathElement.Name + @"\" + "Журнал инцидентов.xlsx"));
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
        public void Dispose(Excel.Application iApp, Excel.Workbook iWb, Tuple<Excel.Worksheet, Excel.Worksheet> nSheets)
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
    }
}