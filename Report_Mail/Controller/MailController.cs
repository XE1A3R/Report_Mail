using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
using Report_Mail.Interface;
using Report_Mail.Model;

namespace Report_Mail.Controller
{
    public class MailController 
    {
	    private readonly IMail _mail;
	    private readonly Label _label1;
	    private List<string> MailSupportError { get; } = new();

	    public MailController(IMail mail, Label label1)
        {
	        _mail = mail;
	        _label1 = label1;
            MailSupportError.AddRange(_mail.MailSupportError);
        }

        public void Send(IEnumerable<Xls> configJsonXls)
        {
	        _label1.Text = @"Отправка SMTP...";
			var smtp = new SmtpClient(_mail.SmtpClient, _mail.Port)
			{
			Credentials = new NetworkCredential(_mail.From, _mail.Password)
			};
			var toAddressListAdd = new MailAddressCollection();
			foreach (var mailAddress in _mail.To.Select(mail => new MailAddress(mail)))
			{
				toAddressListAdd.Add(mailAddress);
			}
			var toAddressListCc = new MailAddressCollection();
			foreach (var mailAddress in _mail.Cc.Select(mail => new MailAddress(mail)))
			{
				toAddressListCc.Add(mailAddress);
			}
			var message = new MailMessage()
			{
				From = new MailAddress(_mail.From, _mail.Name)
			};
			message.To.Add(toAddressListAdd.ToString());
			message.CC.Add(toAddressListCc.ToString());
			message.Subject = _mail.Subject;
			message.Body = _mail.Body;
			foreach (var att in configJsonXls)
			{
				message.Attachments.Add(new Attachment($@"{att.Attachments}\{att.Name}.{att.Format}"));
			}
	        smtp.Send(message);
        }
    }
}