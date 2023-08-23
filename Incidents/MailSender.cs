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
            try
            {
                MailMessage message = new MailMessage();

                foreach(string email in emails)
                {
                    message.To.Add(email);
                }

                message.Subject = mailSubject;
                message.From = new MailAddress("");
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
                SmtpClient smtp = new SmtpClient("");

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
    }
}
