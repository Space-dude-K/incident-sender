using Incidents.Interfaces;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Incidents
{
    class OutlookSender
    {
        private string filePath;
        private string textMessage;
        private string attachmentName;
        private string mailSubject;
        private string recipient;
        private ILogger logger;

        public OutlookSender(string filePath, string textMessage, string attachmentName, string mailSubject, string recipient, ILogger logger)
        {
            this.filePath = filePath;
            this.textMessage = textMessage;
            this.attachmentName = attachmentName;
            this.mailSubject = mailSubject;
            this.recipient = recipient;
            this.logger = logger;
        }
        public void SendEmail()
        {
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();

                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

                // Add the body of the email
                oMsg.HTMLBody = textMessage;

                //Add an attachment.
                String sDisplayName = attachmentName;
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;

                // Attached file
                Outlook.Attachment oAttach = oMsg.Attachments.Add(filePath, iAttachType, iPosition, sDisplayName);

                //Subject line
                oMsg.Subject = mailSubject;

                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;

                // Change the recipient in the next line if necessary.
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(recipient);
                oRecip.Resolve();

                // Send.
                oMsg.Send();

                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
                logger.LogMessage("Daily file SEND. File - " + filePath);
            }
            catch (Exception ex)
            {
                if (logger.GetType() == typeof(LogToTxt))
                {
                    logger.LogMessage("Daily file SEND ERROR. File - " + filePath + " " + ex.Message);
                    Console.WriteLine(ex.Message);
                }
                else
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}
