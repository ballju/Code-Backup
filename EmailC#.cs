using System;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;
using System.ComponentModel;




namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress("tank01@samplecompany.com"); //From Adress 
            msg.To.Add("");//Where to sent can add mutiple addresses
            msg.Subject = "Tank Levels";
            msg.IsBodyHtml = true;
            msg.Body = "Tank # has a level of x% ";//Sample body email message

            SmtpClient smtp = new SmtpClient();
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.UseDefaultCredentials = false; //Set to true if you have authenication
            smtp.EnableSsl = false;
            smtp.Host = "smtp.gmail.com";//gmail's smtp server
            smtp.Port = 25;
            //smtp.Credentials = new NetworkCredential("", ""); //login details if you need them
            smtp.Timeout = int.MaxValue;
            smtp.Send(msg); 
        }
    }
}
