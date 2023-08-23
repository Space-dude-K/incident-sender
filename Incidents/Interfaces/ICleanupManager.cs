using Microsoft.Office.Interop.Excel;
using System;

namespace Incidents.Interfaces
{
    public interface ICleanupManager
    {
        void Cleanup();
        void Dispose(Application iApp, Workbook iWb, Tuple<Worksheet, Worksheet> nSheets);
        void MoveFile(ConnectionManagerPathElement pathElement);
    }
}
