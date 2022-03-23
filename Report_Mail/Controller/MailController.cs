using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
using Report_Mail.Interface;
using Report_Mail.Model;

namespace Report_Mail.Controller
{
    public class MailController : IMail
    {
        private readonly Label _label1;
        public string MailExport { get; set; }
        public string SmtpClient { get; set; }
        public int Port { get; set; }
        public string From { get; set; }
        public string Name { get; set; }
        public string Password { get; set; }
        public List<string> To { get; set; } = new();
        public List<string> Cc { get; set; } = new();
        public string Subject { get; set; }
        public string Body { get; set; }
        public List<string> MailSupportError { get; set; } = new();
        
        public MailController(IMail mail, Label label1)
        {
            _label1 = label1;
            MailExport = mail.MailExport;
            SmtpClient = mail.SmtpClient;
            Port = mail.Port;
            From = mail.From;
            Name = mail.Name;
            Password = mail.Password;
            To.AddRange(mail.To);
            Cc.AddRange(mail.Cc);
            Subject = mail.Subject;
            Body = mail.Body;
            MailSupportError.AddRange(mail.MailSupportError);
        }

        public void Send(IEnumerable<Xls> configJsonXls)
        {
	        _label1.Text = @"Отправка SMTP...";
			var smtp = new SmtpClient(SmtpClient, Port)
			{
			Credentials = new NetworkCredential(From, Password)
			};
			var toAddressListAdd = new MailAddressCollection();
			foreach (var mailAddress in To.Select(mail => new MailAddress(mail)))
			{
				toAddressListAdd.Add(mailAddress);
			}
			var toAddressListCc = new MailAddressCollection();
			foreach (var mailAddress in Cc.Select(mail => new MailAddress(mail)))
			{
				toAddressListCc.Add(mailAddress);
			}
			var message = new MailMessage()
			{
				From = new MailAddress(From, Name)
			};
			message.To.Add(toAddressListAdd.ToString());
			message.CC.Add(toAddressListCc.ToString());
			message.Subject = Subject;
			message.Body = Body;
			foreach (var att in configJsonXls)
			{
				message.Attachments.Add(new Attachment($@"{att.Attachments}\{att.Name}.{att.Format}"));
			}
	        smtp.Send(message);
        }
    }
}