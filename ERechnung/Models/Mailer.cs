using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ERechnung.Models
{
    internal class Mailer
    {
        SmtpClient client;
        public void Setup(string server, int port, bool ssl, string user, string password)
        {
            this.client = new SmtpClient(server, port);
            if (user != ""){
                NetworkCredential credentials = new NetworkCredential(user, password);
                this.client.Credentials = credentials;
            }
        }
        public void SendEmail(string from, string to, string subject, string body, string eRechnungPDF)
        {
            MailAddress fromAddress = new MailAddress(from);
            MailAddress toAddress = new MailAddress(to);
            MailMessage message = new MailMessage(fromAddress, toAddress);
            message.Body = body;
            message.Subject = subject;
            if (IsAttachmentLocked(new FileInfo(eRechnungPDF)))
            {
                Thread.Sleep(2000); // Wait for 2 seconds before retrying
            }
            if (IsAttachmentLocked(new FileInfo(eRechnungPDF)))
            {
                throw new IOException($"The file {eRechnungPDF} is still locked and cannot be attached to the email.");
            }
            message.Attachments.Add(new Attachment(eRechnungPDF, System.Net.Mime.MediaTypeNames.Application.Pdf));
            client.Send(message);
        }

        private bool IsAttachmentLocked(FileInfo file)
        {
            try
            {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    stream.Close();
                }
                return false;

            }
            catch (IOException)
            {
                return true;
            }
        }
    }
}
