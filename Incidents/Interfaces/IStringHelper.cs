using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Incidents.Interfaces
{
    public interface IStringHelper
    {
        bool DateExist(string date, string region);
        Tuple<bool, bool> FileExistAndEmpty(string path);
        string FixedIncidentType(string str);
        bool IsStringsIsEmptyOrNull(string str1, string str2);
        Color SetCellColor(bool state, int counter, bool skip);
        string SetLogString(int counter, string msg, bool skip);
    }
}
