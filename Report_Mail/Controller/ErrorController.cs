using System;
using System.Net.Mail;
using System.Windows.Forms;

namespace Report_Mail.Controller
{
    public class ErrorController
    {
        ErrorController()
        {
					var today = DateTime.Today;
					var day_old = Convert.ToInt32(today.DayOfWeek) + 6;
					var monday_old = today.AddDays(-day_old);
					var sunday_old = monday_old.AddDays(6);
				// 	Log.logger.Error("Error User:{0}, ID:{1}, IP:{2}, Ver:{3} \n" +
				//	"" + Environment.NewLine + "" +
				//	"" + Environment.NewLine + "" +
				//	" {4}" +
					//"" + Environment.NewLine + "" +
				//	" {5}" +
					//"" + Environment.NewLine + "" +
					//"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, ex.Message, ex.StackTrace);

					//var //smtp = new SmtpClient(smtpClient, smtpClient_port)
					//{
					//	Credentials = new NetworkCredential(from_mail, from_Password)
					//};
					var Message = new MailMessage
					{
					//	From = new MailAddress(from_mail, from_mail_name)
					};
					//Message.To.Add(new MailAddress(mail_support_error));
					//Message.Subject = "Report_Mail - Error " + configuration;
					//Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
					//" " + ex.Message + "";
					//Message.Attachments.Add(new Attachment("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"));
					try
					{
					//	smtp.Send(Message);
					}
					catch (SmtpException)
					{
						MessageBox.Show("Ошибка!", "smtp");
					}
        }
    }
}