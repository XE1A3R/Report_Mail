using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Net;
using System.IO;

namespace Report_Mail
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
			label1.Text = "";
			//Thread Run = new Thread(Worker_1);
			//Run.Start();
			Worker_1();
		}

		void Worker_1()
		{
			progressBar1.Visible = true;
			backgroundWorker1.RunWorkerAsync();
		}

		void Worker_2()
		{
			backgroundWorker2.RunWorkerAsync();
		}

		private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
			this.Invoke(new Action(Report_1));
		}

		private void BackgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
		{

			//this.Invoke(new Action(Report_2));
		}
		private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			progressBar1.Value = e.ProgressPercentage;
		}

		private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			Invoke(new Action(Sleep_Exit));
		}
		void Report_1()
		{
			OdbcConnection cnS11 = new OdbcConnection(Properties.Settings.Default.S11);
			OdbcDataAdapter adapter1 = new OdbcDataAdapter("SELECT CONCAT('с ', DATE_FORMAT(DATE_sub(CURDATE(), INTERVAL 18 HOUR), '%d.%m.%Y %H:%i'), ' по ', DATE_FORMAT(DATE_ADD(CURDATE(), INTERVAL 6 HOUR), '%d.%m.%Y %H:%i'))", cnS11);
			DataTable table1 = new DataTable();
			adapter1.Fill(table1);
			dataGridView2.DataSource = table1;
			var xls = dataGridView2[0, 0].Value.ToString();
			label1.Text = "Выполняется процедура...";
			OdbcDataAdapter adapter = new OdbcDataAdapter("CALL procStatementForPastDay", cnS11);
			DataTable table = new DataTable();
			adapter.Fill(table);
			dataGridView1.DataSource = table;
			Console.WriteLine(xls);
			progressBar1.Maximum = dataGridView1.RowCount + dataGridView1.ColumnCount;
			backgroundWorker1.ReportProgress(dataGridView1.RowCount + dataGridView1.ColumnCount);
			backgroundWorker1.ReportProgress(0);
			label1.Text = "Выгрузка в EXCEL...";
			Excel.Application xlApp;
			Excel.Workbook xlWorkBook;
			Excel.Worksheet xlWorkSheet;
			object misValue = System.Reflection.Missing.Value;

			Int16 i, j;
			xlApp = new Excel.Application();
			xlWorkBook = xlApp.Workbooks.Add(misValue);
			xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

			xlApp.Cells[1, 1] = "Отчет";
			xlApp.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlRight;
			xlApp.Cells[1, 2] = xls;
			xlApp.Cells[2, 1] = "Выписан из Отделения";
			xlApp.Cells[2, 2] = "Врач";
			xlApp.Cells[2, 3] = "Выписано";
			xlApp.Cells[2, 4] = "Без диагноза";
			xlWorkSheet.Cells[1, 4].EntireRow.Font.Bold = xlWorkSheet.Cells[2, 4].EntireRow.Font.Bold = true;


			int h = 1;
			//xlApp.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
			//xlApp.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
			//xlApp.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
			for (i = 0; i < dataGridView1.RowCount; i++)
			{
				h++;				
				for (j = 0; j < dataGridView1.ColumnCount; j++)
				{					
					backgroundWorker1.ReportProgress(i + j);
					xlApp.Cells[h + 1, 4].Interior.Color = Color.Red;
					xlWorkSheet.Cells[h + 1, j + 1] = dataGridView1[j, i].Value.ToString();
				}
			}
			backgroundWorker1.ReportProgress(dataGridView1.ColumnCount + dataGridView1.RowCount);


			((Excel.Range)xlWorkSheet.Columns[1]).AutoFit();
			((Excel.Range)xlWorkSheet.Columns[2]).AutoFit();
			((Excel.Range)xlWorkSheet.Columns[3]).AutoFit();
			((Excel.Range)xlWorkSheet.Columns[4]).AutoFit();

			xlWorkBook.SaveAs(@"" + Environment.CurrentDirectory + "/Test.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
			xlWorkBook.Close(true, misValue, misValue);
			xlApp.Quit();

			ReleaseObject(xlWorkSheet);
			ReleaseObject(xlWorkBook);
			ReleaseObject(xlApp);
			try
			{
				//	SmtpClient smtp = new SmtpClient("mail.gb15.ru", 25)
				//	{
				//		Credentials = new NetworkCredential("robot@gb15.ru", "1oc@1RoBoT")
				//	};
				//	MailMessage Message = new MailMessage
				//	{
				//		From = new MailAddress("robot@gb15.ru", "Report")
				//	};
				//	Message.To.Add(new MailAddress(Properties.Settings.Default.mail_add));
				//	//Message.CC.Add(new MailAddress(Properties.Settings.Default.mail_cc));
				//	Message.Subject = xls;
				//	Message.Body = "Добрый день. \n" +
				//					"" + Environment.NewLine + "" +
				//					"Отчет во вложении." +
				//					"" + Environment.NewLine + "" +
				//					"" + Environment.NewLine + "" +
				//					"--\n" +
				//					"С уважением,\n" +
				//					"Группа МИС\n" +
				//					"СПб ГБУЗ 'Городская больница № 15'";
				//	Message.Attachments.Add(new Attachment("" + Environment.CurrentDirectory + "/Test.xls"));
				//	try
				//	{

				//		smtp.Send(Message);
				//		MessageBox.Show("Отправлено.", "smtp");
				//	}
				//	catch (SmtpException)
				//	{					

				//		MessageBox.Show("Ошибка!", "smtp");
				//	}
				//	Message.Attachments.Dispose();
				//File.Delete(@"" + Environment.CurrentDirectory + "/Test.xls");

			}
			catch (Exception ex)
			{
				//	Log.logger.Error("Error User:{0}, ID:{1}, IP:{2}, Ver:{3} \n" +
				//		"" + Environment.NewLine + "" +
				//		"" + Environment.NewLine + "" +
				//		"     {4}" +
				//		"     {5}" +
				//		"" + Environment.NewLine + "" +
				//		"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, отчетToolStripMenuItem, ex.Message);
				//	if (выводОшибокToolStripMenuItem.Checked == true)
				//	{
				//		MessageBox.Show(ex.Message, "Error");
				//	}
				//	SmtpClient smtp = new SmtpClient("mail.gb15.ru", 25)
				//	{
				//		Credentials = new NetworkCredential("robot@gb15.ru", "1oc@1RoBoT")
				//	};
				//	MailMessage Message = new MailMessage
				//	{
				//		From = new MailAddress("robot@gb15.ru", "Report")
				//	};
				//	Message.To.Add(new MailAddress(Properties.Settings.Default.mail_Error));
				//	Message.Subject = "Error";
				//	Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
				//		"     " + отчетToolStripMenuItem + " \n" +
				//		"     " + ex.Message + "";
				//	Message.Attachments.Add(new Attachment("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"));
				//	try
				//	{
				//		smtp.Send(Message);
				//	}
				//	catch (SmtpException)
				//	{
				//		if (выводОшибокToolStripMenuItem.Checked == true)
				//		{
				//			MessageBox.Show("Ошибка!", "smtp");
				//		}

				//	}
			}
			//progressBar1.Visible = false;
		}

		void Report_2()
		{

		}

		void Sleep_Exit()
		{
			timer1.Enabled = true;
			int times = timer1.Interval;
			notifyIcon1.BalloonTipText = "Завершение программы начнется через " + Convert.ToString(times / 1000) + " екунд";
			notifyIcon1.ShowBalloonTip(5000);

		}
		private void ReleaseObject(object obj)
		{
			try
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
			}
			catch (Exception ex)
			{
				obj = null;
				MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
			}
			finally
			{
				GC.Collect();
			}
		}

		private void Timer1_Tick(object sender, EventArgs e)
		{
			Application.Exit();
		}
	}
}
