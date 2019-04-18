using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Text;

namespace IRISTicketTracker
{
    public class MailUtility
    {
        public void SendMail(string from, string to, string subject, string body)
        {
            MailMessage mail = new MailMessage(from, to);
            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = true;
            client.Host = "smtp.na.jnj.com";
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = true;
            client.Send(mail);
        }
    }
}
