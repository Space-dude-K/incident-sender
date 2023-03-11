using Incidents.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;

namespace Incidents
{
    class MailSender : IMailSender
    {
        private readonly ILogger logger;
        public MailSender(ILogger logger)
        {
            this.logger = logger;
        }
        public void SendEmail(string filePath, string textMessage, string mailSubject, List<string> emails)
        {
            //  string filePath, string textMessage, string attachmentName, string mailSubject, string recipient, ILogger logger
            //string recipient1 = @"G600-U115@mfrb.by";
            //string recipient2 = @"dmitriy.lukashkov@mogrfo.mogfu.gov.by";

            try
            {
                MailMessage message = new MailMessage();

                foreach(string email in emails)
                {
                    message.To.Add(email);
                }

                message.Subject = mailSubject;
                message.From = new MailAddress("G600-Mailer@mfrb.by");
                message.Body = textMessage;

                // Create the file attachment for this e-mail message.
                Attachment data = new Attachment(filePath, MediaTypeNames.Application.Octet);

                // Add time stamp information for the file.
                ContentDisposition disposition = data.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(filePath);
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(filePath);
                disposition.ReadDate = System.IO.File.GetLastAccessTime(filePath);

                // Add the file attachment to this e-mail message.
                message.Attachments.Add(data);

                //SmtpClient smtp = new SmtpClient("gkdoc.gkmogilev.minfin.by");
                SmtpClient smtp = new SmtpClient("M000-MBX1.mfrb.by");

                foreach (string email in emails)
                {
                    logger.LogMessage("Sending msg to -> " + email);
                }

                smtp.Send(message);
            }
            catch (Exception ex)
            {
                logger.LogMessage(ex.Message);
            }
        }
        /// <summary>
        /// Компоновщик. Создает xml-шаблон с таблицей. Create xml-template for table.
        /// </summary>
        /// <param name="msgData">Подготовленные данные. Data: DNS name - ip - shutdown delay - domain time</param>
        /// <returns></returns>
        private void MailComposer(List<string> tableData, string domainTime)
        {
            /*
            StringBuilder result = new StringBuilder();

            // Table 1
            result.Append("<head><style>table, td, th { border: 1px solid black; width: 710px; }</style></head>");
            result.Append("<h1>Smdo zero file checker report.</h1>");
            result.Append("<h2>Domain time: " + domainTime + "</h2>");
            result.Append("<table><tr><th>File</th><th>Folder</th></tr>");

            foreach (string data in tableData)
            {
                result.Append(
                    "<tr>" +
                    "<td style='text-align:center'>" + System.IO.Path.GetFileName(data) + "</td>" +
                    "<td style='text-align:center'>" + "<a href=" + System.IO.Path.GetDirectoryName(data) + ">Folder link</a>" + "</td>" +
                    "</tr>");
            }

            result.Append("</table>");
            result.Append("<h3><a href=" + logger.LoggerPath + ">Log file link</a></h3>");
            result.Append("<h4><a href=" + confPath + ">Edit settings</a></h4>");

            return result.ToString();
             */
        }
    }
}
