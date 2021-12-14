using System;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;

namespace Report_Mail
{
	public partial class Form1 : Form
	{
		private int _x;

		private readonly ConfigApp _config;
		

		[Obsolete]
		public Form1(string[] file)
		{
			InitializeComponent();
			_config = new ConfigApp(file);
		}

		void Worker_1()
		{
			label1.Text = "";
			backgroundWorker1.RunWorkerAsync();
		}

		private void Label()
		{
			progressBar1.Visible = true;
			progressBar1.Maximum = dataGridView1.RowCount + dataGridView1.ColumnCount;

			if (_x == 1)
			{
				label1.Text = @"Выполняется процедура...";
				progressBar1.Style = ProgressBarStyle.Marquee;
			}
			else if (_x == 2)
			{
				label1.Text = @"Создание нового файла EXCEL...";
				progressBar1.Style = ProgressBarStyle.Marquee;
			}
			else if (_x == 3)
			{
				label1.Text = @"Выгрузка в EXCEL... ";
				progressBar1.Style = ProgressBarStyle.Continuous;
			}
			else if (_x == 4)
			{
				label1.Text = @"Сохранение EXCEL...";
				progressBar1.Style = ProgressBarStyle.Marquee;
			}
			else if (_x == 5)
			{
				label1.Text = @"Отправка SMTP...";
				progressBar1.Style = ProgressBarStyle.Marquee;
			}
			else if (_x == 6)
			{
				label1.Text = @"Отправлено.";
				progressBar1.Style = ProgressBarStyle.Continuous;
			}
			else if (_x == 7)
			{
				label1.Text = @"Ошибка.";
				progressBar1.Style = ProgressBarStyle.Continuous;
			}
			else if (_x == 8)
			{
				label1.Text = @"Удаление временных файлов...";
				progressBar1.Style = ProgressBarStyle.Marquee;
			}
			else if (_x == 9)
			{
				label1.Text = @"Выполнено.";
				progressBar1.Style = ProgressBarStyle.Continuous;
			}
		}

		[Obsolete]
		private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
			_x = 2;
			backgroundWorker1.ReportProgress(0);
			Invoke(new Action(Label));
			var excel = new Excel();
			if (_config.ConfigJson == null) return;
			foreach (var xls1 in _config.ConfigJson.Xls)
			{
				foreach (var sheet in xls1.Sheets)
				{
					_x = 2;
					Invoke(new Action(Label));
					excel.CreateSheet(sheet.Name);
					foreach (var location in sheet.Locations)
					{
						_x = 1;
						Invoke(new Action(Label));
						Excel.CreateTable(location.Request, ref dataGridView1);
						_x = 3;
						Invoke(new Action(Label));
						if (excel.ExcelSheet != null)
							excel.ExcelSheet[excel.ExcelSheet.Count - 1].InsertData(location.Column, location.Row,
								dataGridView1, backgroundWorker1, excel.ExcelSheet);
					}
				}

				_x = 4;
				Invoke(new Action(Label));
				excel.Save(@$"{xls1.Attachments}\{xls1.name}.xls");
				foreach (var item in _config.ConfigJson.Mail)
				{
					_x = 5;
					Invoke(new Action(Label));
					SmtpClient smtp = new SmtpClient(item.SmtpClient, item.Port)
					{
						Credentials = new NetworkCredential(item.From, item.Password)
					};
					var toAddressListAdd = new MailAddressCollection();
					foreach (var mailAddress in item.To.Select(mail => new MailAddress(mail)))
					{
						toAddressListAdd.Add(mailAddress);
					}
					MailAddressCollection toAddressListCc = new MailAddressCollection();
					foreach (var mailAddress in item.Cc.Select(mail => new MailAddress(mail)))
					{
						toAddressListCc.Add(mailAddress);
					}
					var message = new MailMessage()
					{
						From = new MailAddress(item.From, item.Name)
					};
					message.To.Add(toAddressListAdd.ToString());
					message.CC.Add(toAddressListCc.ToString());
					message.Subject = item.Subject;
					message.Body = item.Body;
					foreach (var att in _config.ConfigJson.Xls)
					{
						message.Attachments.Add(new Attachment($@"{att.Attachments}\{att.name}.xls"));
					}
					smtp.Send(message);
					backgroundWorker1.ReportProgress(dataGridView1.ColumnCount + dataGridView1.RowCount);
					_x = 6;
					Invoke(new Action(Label));
				}
				
			}
		}

		private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			progressBar1.Value = e.ProgressPercentage;
		}

		private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			Invoke(new Action(Sleep_Exit));
		}

		void Sleep_Exit()
		{
			timer1.Enabled = true;
			int times = timer1.Interval;
			notifyIcon1.BalloonTipText = @"Завершение программы начнется через " + Convert.ToString(times / 1000) + @" секунд";
			notifyIcon1.ShowBalloonTip(5000);
			Properties.Settings.Default.Save();
		}

		private void Timer1_Tick(object sender, EventArgs e)
		{
			Application.Exit();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			Worker_1();
		}
	}
}
