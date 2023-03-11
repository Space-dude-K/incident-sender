using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Incidents.Interfaces
{
    internal interface IMailSender
    {
        void SendEmail(string filePath, string textMessage, string mailSubject, List<string> emails);
    }
}
