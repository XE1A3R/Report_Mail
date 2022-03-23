using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceModel.Dispatcher;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using Report_Mail.Controller;
using Report_Mail.Model;

namespace Report_Mail
{
	public partial class Form1 : Form
	{
		private readonly ConfigAppController _config;
		
		[Obsolete]
		public Form1(IReadOnlyList<string> file)
		{
			InitializeComponent();
			_config = new ConfigAppController(file);
		}

		private void Worker_1()
		{
			label1.Text = "";
			backgroundWorker1.RunWorkerAsync();
		}

		public void ProgressValueMax(int value)
		{
			progressBar1.Maximum = value;
		}
		private void Label()
		{
			progressBar1.Visible = true;
			progressBar1.Maximum = 100;
		}

		[Obsolete]
		private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
			try
			{
				backgroundWorker1.ReportProgress(0);
				Invoke(new Action(Label));
				if (_config.ConfigJson == null) return;
				var excel = _config.ConfigJson.Xls.Select(xls => new ExcelWindowController(xls, label1))
					.ToList();

				foreach (var item in excel)
				{
					item.CreateSheet();
					item.Save();
				}

				var mails = _config.ConfigJson.Mail.Select(mail => new MailController(mail, label1)).ToList();
				foreach (var mail in mails)
				{
					mail.Send(_config.ConfigJson.Xls);
				}
			}
			catch (Exception exception)
			{
				label1.Text = @$"Ошибка.\\n{exception.Message}";
				Console.WriteLine(exception.Message);
				throw;
			}
			label1.Text = @"Выполнено.";
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
