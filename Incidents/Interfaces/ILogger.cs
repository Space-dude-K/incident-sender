using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Incidents.Interfaces
{
    public interface ILogger
    {
        void LogMessage(string msg, bool isEndLine = false);
    }
}
