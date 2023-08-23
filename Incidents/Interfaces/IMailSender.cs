using System.Collections.Generic;

namespace Incidents.Interfaces
{
    internal interface IMailSender
    {
        void SendEmail(string filePath, string textMessage, string mailSubject, List<string> emails);
    }
}
