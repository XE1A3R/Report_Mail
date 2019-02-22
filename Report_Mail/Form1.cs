using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Net;
using System.IO;
using MySql.Data.MySqlClient;
using Prof;
using System.Configuration;

namespace Report_Mail
{
	public partial class Form1 : Form
	{
		public int x = 0;
		public bool Sel_1_on_off_DataTime;
		public string Sel_Request_DataTime;
		public bool Sel_2_on_off;
		public string Sel_2_request;
		public string Sel_3_request;
		public string Sel_4_request;
		public string xlApp_Cells_1_1;
		public string xlApp_Cells_1_1_HorizontalAlignment;
		public string xlApp_Cells_2_1;
		public string xlApp_Cells_1_2_HorizontalAlignment;
		public string xlApp_Cells_2_2;
		public string xlApp_Cells_2_3;
		public string xlApp_Cells_2_4;
		public bool xlApp_Cells_1_4_2_4_Font;
		public int int_h;
		public string xlWorkBook_SaveAs;
		public string smtpClient;
		public int smtpClient_port;
		public string from_mail;
		public string from_mail_name;
		public string from_Password;
		public string mail_to;
		public string mail_cc;
		public string subject;
		public string Body;
		public int attachments;
		public string attachments1;
		public string attachments2;
		public string attachments3;
		public string attachments4;
		public string mail_support_error;
		public string xls;
		public int Do =1;
		MySqlConnection cnS11 = SqlConn.DBUtilsS11.GetDBConnection();
		public Form1(string[] file)
		{
			if (file.Length > 0)
			{
				if (File.Exists(@"C:\confs\" + file[1] + ".config"))
				{
					try
					{
						ExeConfigurationFileMap configFile = new ExeConfigurationFileMap
						{
							ExeConfigFilename = Path.Combine(@"C:\confs\", file[1] + ".config")
						};
						Configuration currentConfiguration = ConfigurationManager.OpenMappedExeConfiguration(configFile, ConfigurationUserLevel.None);
						Sel_1_on_off_DataTime = bool.Parse(currentConfiguration.AppSettings.Settings["Sel_1_on_off_DataTime"].Value);
						Sel_Request_DataTime = currentConfiguration.AppSettings.Settings["Sel_Request_DataTime_xlApp.Cells[1, 2]"].Value;
						Sel_2_on_off = bool.Parse(currentConfiguration.AppSettings.Settings["sel_2_on_off"].Value);
						Sel_2_request = currentConfiguration.AppSettings.Settings["Sel_2_request"].Value;
						Sel_3_request = currentConfiguration.AppSettings.Settings["Sel_3_request"].Value;
						Sel_4_request = currentConfiguration.AppSettings.Settings["Sel_4_request"].Value;
						xlApp_Cells_1_1 = currentConfiguration.AppSettings.Settings["xlApp.Cells[1, 1]"].Value;
						xlApp_Cells_1_1_HorizontalAlignment = currentConfiguration.AppSettings.Settings["xlApp.Cells[1, 1].HorizontalAlignment"].Value;
						xlApp_Cells_2_1 = currentConfiguration.AppSettings.Settings["xlApp.Cells[2, 1]"].Value;
						xlApp_Cells_1_2_HorizontalAlignment = currentConfiguration.AppSettings.Settings["xlApp.Cells[1, 2].HorizontalAlignment"].Value;
						xlApp_Cells_2_2 = currentConfiguration.AppSettings.Settings["xlApp.Cells[2, 2]"].Value;
						xlApp_Cells_2_3 = currentConfiguration.AppSettings.Settings["xlApp.Cells[2, 3]"].Value;
						xlApp_Cells_2_4 = currentConfiguration.AppSettings.Settings["xlApp.Cells[2, 4]"].Value;
						xlApp_Cells_1_4_2_4_Font = bool.Parse(currentConfiguration.AppSettings.Settings["xlWorkSheet.Cells[1, 4].EntireRow.Font.Bold = xlWorkSheet.Cells[2, 4].EntireRow.Font.Bold"].Value);
						int_h = int.Parse(currentConfiguration.AppSettings.Settings["int h"].Value);
						xlWorkBook_SaveAs = currentConfiguration.AppSettings.Settings["xlWorkBook.SaveAs"].Value;
						smtpClient = currentConfiguration.AppSettings.Settings["smtpClient"].Value;
						smtpClient_port = int.Parse(currentConfiguration.AppSettings.Settings["smtpClient_port"].Value);
						from_mail = currentConfiguration.AppSettings.Settings["from_mail"].Value;
						from_mail_name = currentConfiguration.AppSettings.Settings["from_mail_name"].Value;
						from_Password = currentConfiguration.AppSettings.Settings["from_Password"].Value;
						mail_to = currentConfiguration.AppSettings.Settings["mail_to"].Value;
						mail_cc = currentConfiguration.AppSettings.Settings["mail_cc"].Value;
						subject = currentConfiguration.AppSettings.Settings["subject"].Value;
						Body = currentConfiguration.AppSettings.Settings["Body"].Value;
						attachments = int.Parse(currentConfiguration.AppSettings.Settings["attachments"].Value);
						attachments1 = currentConfiguration.AppSettings.Settings["attachments1"].Value;
						attachments2 = currentConfiguration.AppSettings.Settings["attachments2"].Value;
						attachments3 = currentConfiguration.AppSettings.Settings["attachments3"].Value;
						attachments4 = currentConfiguration.AppSettings.Settings["attachments4"].Value;
						mail_support_error = currentConfiguration.AppSettings.Settings["mail_support_error"].Value;
						InitializeComponent();
					}
					catch (Exception ex)
					{
						MessageBox.Show("Error - " + ex.Message);
					}
				}
				else
				{
					MessageBox.Show("Файл " + file[1] + ".config ненайден");
					Application.Exit();
				}
			}
		}

		void Worker_1()
		{
			label1.Text = "";
			backgroundWorker1.RunWorkerAsync();
		}

		void Label()
		{
			progressBar1.Visible = true;
			progressBar1.Maximum = dataGridView1.RowCount + dataGridView1.ColumnCount;
			if (x == 1)
				label1.Text = "Выполняется процедура...";
			else if (x == 2)
				label1.Text = "Создание нового файла EXCEL...";
			else if (x == 3)
				label1.Text = "Выгрузка в EXCEL... " + xlWorkBook_SaveAs;
			else if (x == 4)
				label1.Text = "Сохранение EXCEL...";
			else if (x == 5)
				label1.Text = "Отправка SMTP...";
			else if (x == 6)
				label1.Text = "Отправлено.";
			else if (x == 7)
				label1.Text = "Ошибка.";
			else if (x == 8)
				label1.Text = "Удаление временных файлов...";
			else if (x == 9)
				label1.Text = "Выполнено.";
		}
		
		[Obsolete]
		private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
			try
			{
				if (Sel_1_on_off_DataTime)
				{
					MySqlDataAdapter adapter1 = new MySqlDataAdapter(Sel_Request_DataTime, cnS11);
					//MySqlDataAdapter adapter1 = new MySqlDataAdapter("SELECT CONCAT('с ', DATE_FORMAT(DATE_sub(CURDATE(), INTERVAL 18 HOUR), '%d.%m.%Y %H:%i'), ' по ', DATE_FORMAT(DATE_ADD(CURDATE(), INTERVAL 6 HOUR), '%d.%m.%Y %H:%i'))", cnS11);
					DataTable table1 = new DataTable();
					adapter1.Fill(table1);
					dataGridView2.DataSource = table1;
					xls = dataGridView2[0, 0].Value.ToString();
				}
				do
				{
					if (Sel_2_on_off)
					{
						
						if (Do == 2)
							Sel_2_request = Sel_3_request;
						else if (Do == 3)
							Sel_2_request = Sel_4_request;
						x = 1;
						Invoke(new Action(Label));
						MySqlDataAdapter adapter = new MySqlDataAdapter(Sel_2_request, cnS11);
						DataTable table = new DataTable();
						adapter.Fill(table);
						dataGridView1.DataSource = table;
						Invoke(new Action(Label));
						//progressBar1.Maximum = dataGridView1.RowCount + dataGridView1.ColumnCount;
						backgroundWorker1.ReportProgress(dataGridView1.RowCount + dataGridView1.ColumnCount);
						backgroundWorker1.ReportProgress(0);
					}

					if (Do == 2)
						xlWorkBook_SaveAs=attachments2;
					if (Do == 3)
						xlWorkBook_SaveAs = attachments3;
					Excel.Application xlApp;
					Excel.Workbook xlWorkBook;
					Excel.Worksheet xlWorkSheet;
					object misValue = System.Reflection.Missing.Value;
					x = 2;
					Invoke(new Action(Label));
					//label1.Text = "Создание нового файла EXCEL...";
					Int16 i, j;
					int h = int_h;
					xlApp = new Excel.Application();
					xlWorkBook = xlApp.Workbooks.Add(misValue);
					xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

					xlApp.Cells[1, 1] = xlApp_Cells_1_1;
					xlApp.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlRight;
					xlApp.Cells[1, 2] = xls;
					xlApp.Cells[2, 1] = xlApp_Cells_2_1;
					xlApp.Cells[2, 2] = xlApp_Cells_2_2;
					xlApp.Cells[2, 3] = xlApp_Cells_2_3;
					xlApp.Cells[2, 4] = xlApp_Cells_2_4;
					xlWorkSheet.Cells[1, 4].EntireRow.Font.Bold = xlWorkSheet.Cells[2, 4].EntireRow.Font.Bold = xlApp_Cells_1_4_2_4_Font;
					//xlApp.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
					//xlApp.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
					//xlApp.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
					x = 3;
					Invoke(new Action(Label));
					//label1.Text = "Выгрузка в EXCEL...";
					for (i = 0; i < dataGridView1.RowCount; i++)
					{
						h++;
						for (j = 0; j < dataGridView1.ColumnCount; j++)
						{
							backgroundWorker1.ReportProgress(i + j);
							xlApp.Cells[h + 1, 4].Font.Color = Color.Red;
							xlWorkSheet.Cells[h + 1, j + 1] = dataGridView1[j, i].Value.ToString();
						}
					}
					//backgroundWorker1.ReportProgress(dataGridView1.ColumnCount + dataGridView1.RowCount);


					((Excel.Range)xlWorkSheet.Columns[1]).AutoFit();
					((Excel.Range)xlWorkSheet.Columns[2]).AutoFit();
					((Excel.Range)xlWorkSheet.Columns[3]).AutoFit();
					((Excel.Range)xlWorkSheet.Columns[4]).AutoFit();
					x = 4;
					Invoke(new Action(Label));
					//label1.Text = "Сохранение EXCEL...";
					xlWorkBook.SaveAs(xlWorkBook_SaveAs, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
					xlWorkBook.Close(true, misValue, misValue);
					xlApp.Quit();

					ReleaseObject(xlWorkSheet);
					ReleaseObject(xlWorkBook);
					ReleaseObject(xlApp);
					
					if (Do == attachments)
						break;
					Do++;
				}				
				while (attachments > 1);
				try
				{

					x = 5;
					Invoke(new Action(Label));
					//label1.Text = "Отправка SMTP...";
					SmtpClient smtp = new SmtpClient(smtpClient, smtpClient_port)
					{
						Credentials = new NetworkCredential(from_mail, from_Password)
					};
					MailAddressCollection TO_addressList_add = new MailAddressCollection();
					foreach (var mail in mail_to.Split(','))
					{
						MailAddress mailAddress = new MailAddress(mail);
						TO_addressList_add.Add(mailAddress);
					}
					MailAddressCollection TO_addressList_cc = new MailAddressCollection();
					foreach (var mail in mail_cc.Split(','))
					{
						MailAddress mailAddress = new MailAddress(mail);
						TO_addressList_cc.Add(mailAddress);
					}
					MailMessage Message = new MailMessage()
					{
						From = new MailAddress(from_mail, from_mail_name)
					};
					Message.To.Add(TO_addressList_add.ToString());
					Message.CC.Add(TO_addressList_cc.ToString());
					Message.Subject = subject + xls;
					Message.Body = Body;
					switch (Do)
					{
						case 0:

							break;
						case 1:
							Message.Attachments.Add(new Attachment(xlWorkBook_SaveAs));
							break;
						case 2:
							Message.Attachments.Add(new Attachment(attachments1));
							Message.Attachments.Add(new Attachment(attachments2));
							break;
						case 3:
							Message.Attachments.Add(new Attachment(attachments1));
							Message.Attachments.Add(new Attachment(attachments2));
							Message.Attachments.Add(new Attachment(attachments3));
							break;
						case 4:
							Message.Attachments.Add(new Attachment(attachments1));
							Message.Attachments.Add(new Attachment(attachments2));
							Message.Attachments.Add(new Attachment(attachments3));
							Message.Attachments.Add(new Attachment(attachments4));
							break;
					}
					try
					{

						smtp.Send(Message);
						backgroundWorker1.ReportProgress(dataGridView1.ColumnCount + dataGridView1.RowCount);
						x = 6;
						Invoke(new Action(Label));
						//label1.Text = "Выполнено.";
					}
					catch (SmtpException)
					{
						x = 7;
						Invoke(new Action(Label));
						backgroundWorker1.CancelAsync();
						//label1.Text = "Ошибка.";
					}
					Message.Attachments.Dispose();
					x = 8;
					Invoke(new Action(Label));
					x = 9;
					Invoke(new Action(Label));
					if(attachments==1)
						File.Delete(xlWorkBook_SaveAs);
					else if (attachments == 2)
					{
						File.Delete(attachments1);
						File.Delete(attachments2);
					}
					else if (attachments == 3)
					{
						File.Delete(attachments1);
						File.Delete(attachments2);
						File.Delete(attachments3);
					}
					else if (attachments == 4)
					{
						File.Delete(attachments1);
						File.Delete(attachments2);
						File.Delete(attachments3);
						File.Delete(attachments4);
					}

				}
				catch (Exception ex)
				{
					var today = DateTime.Today;
					var day_old = Convert.ToInt32(today.DayOfWeek) + 6;
					var monday_old = today.AddDays(-day_old);
					var sunday_old = monday_old.AddDays(6);
					Log.logger.Error("Error User:{0}, ID:{1}, IP:{2}, Ver:{3} \n" +
							"" + Environment.NewLine + "" +
							"" + Environment.NewLine + "" +
							"     {4}" +
							"" + Environment.NewLine + "" +
							"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, ex.Message);

					SmtpClient smtp = new SmtpClient(smtpClient, smtpClient_port)
					{
						Credentials = new NetworkCredential(from_mail, from_Password)
					};
					MailMessage Message = new MailMessage
					{
						From = new MailAddress(from_mail, from_mail_name)
					};
					Message.To.Add(new MailAddress(mail_support_error));
					Message.Subject = "Error";
					Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
						"     " + ex.Message + "";
					Message.Attachments.Add(new Attachment("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"));
					try
					{
						smtp.Send(Message);
					}
					catch (SmtpException)
					{
						MessageBox.Show("Ошибка!", "smtp");
					}
				}
			}
			catch (Exception ex)
			{
				var today = DateTime.Today;
				var day_old = Convert.ToInt32(today.DayOfWeek) + 6;
				var monday_old = today.AddDays(-day_old);
				var sunday_old = monday_old.AddDays(6);
				Log.logger.Error("Error User:{0}, ID:{1}, IP:{2}, Ver:{3} \n" +
						"" + Environment.NewLine + "" +
						"" + Environment.NewLine + "" +
						"     {4}" +
						"" + Environment.NewLine + "" +
						"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, ex.Message);

				SmtpClient smtp = new SmtpClient(smtpClient, smtpClient_port)
				{
					Credentials = new NetworkCredential(from_mail, from_Password)
				};
				MailMessage Message = new MailMessage
				{
					From = new MailAddress(from_mail, from_mail_name)
				};
				Message.To.Add(new MailAddress(mail_support_error));
				Message.Subject = "Error";
				Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
					"     " + ex.Message + "";
				Message.Attachments.Add(new Attachment("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"));
				try
				{
					smtp.Send(Message);
				}
				catch (SmtpException)
				{
					MessageBox.Show("Ошибка!", "smtp");
				}
			}
			//this.Invoke(new Action(Report_1));
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
			notifyIcon1.BalloonTipText = "Завершение программы начнется через " + Convert.ToString(times / 1000) + " секунд";
			notifyIcon1.ShowBalloonTip(5000);
			Properties.Settings.Default.Save();

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
		
		private void Form1_Load(object sender, EventArgs e)
		{
			Worker_1();
		}
	}
}
