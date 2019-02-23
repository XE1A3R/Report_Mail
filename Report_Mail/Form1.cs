using Prof;
using System;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Report_Mail
{
	public partial class Form1 : Form
	{
		private int x = 0;
		private bool Sel_DataTime;
		private string Sel_Request_DataTime;
		private bool Sel_2;
		private string Sel_2_request;
		private string Sel_3_request;
		private string Sel_4_request;
		private bool excel_export;
		private string conf1;
		private string conf2;
		private string conf3;
		private string conf4;
		private string xlApp_Cells_1_1;
		private string xlApp_Cells_1_1_HorizontalAlignment;
		private string xlApp_Cells_2_1;
		private string xlApp_Cells_1_2_HorizontalAlignment;
		private string xlApp_Cells_2_2;
		private string xlApp_Cells_2_3;
		private string xlApp_Cells_2_4;
		private bool xlApp_Cells_1_4_2_4_Font;
		private int int_h;
		private string xlWorkBook_SaveAs;
		private bool mail_export;
		private string smtpClient;
		private int smtpClient_port;
		private string from_mail;
		private string from_mail_name;
		private string from_Password;
		private string mail_to;
		private string mail_cc;
		private string subject;
		private string Body;
		private int attachments;
		private string attachments1;
		private string attachments2;
		private string attachments3;
		private string attachments4;
		private string mail_support_error;
		private string xls;
		private int Do = 1;
		readonly OdbcConnection cnS11 = new OdbcConnection(Properties.Settings.Default.S11);

		[Obsolete]
		public Form1(string[] file)
		{
			InitializeComponent();
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
						Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFile, ConfigurationUserLevel.None);
						Sel_DataTime = bool.Parse(config.AppSettings.Settings["Sel_DataTime"].Value);
						Sel_Request_DataTime = config.AppSettings.Settings["Sel_Request_DataTime"].Value;
						Sel_2 = bool.Parse(config.AppSettings.Settings["sel_2"].Value);
						Sel_2_request = config.AppSettings.Settings["Sel_2_request"].Value;
						Sel_3_request = config.AppSettings.Settings["Sel_3_request"].Value;
						Sel_4_request = config.AppSettings.Settings["Sel_4_request"].Value;
						excel_export = bool.Parse(config.AppSettings.Settings["excel_export"].Value);
						conf1 = config.AppSettings.Settings["conf1"].Value;
						conf2 = config.AppSettings.Settings["conf2"].Value;
						conf3 = config.AppSettings.Settings["conf3"].Value;
						conf4 = config.AppSettings.Settings["conf4"].Value;
						smtpClient = config.AppSettings.Settings["smtpClient"].Value;
						smtpClient_port = int.Parse(config.AppSettings.Settings["smtpClient_port"].Value);
						from_mail = config.AppSettings.Settings["from_mail"].Value;
						from_mail_name = config.AppSettings.Settings["from_mail_name"].Value;
						from_Password = config.AppSettings.Settings["from_Password"].Value;
						mail_to = config.AppSettings.Settings["mail_to"].Value;
						mail_cc = config.AppSettings.Settings["mail_cc"].Value;
						subject = config.AppSettings.Settings["subject"].Value;
						Body = config.AppSettings.Settings["Body"].Value;
						attachments = int.Parse(config.AppSettings.Settings["attachments"].Value);
						attachments1 = config.AppSettings.Settings["attachments1"].Value;
						attachments2 = config.AppSettings.Settings["attachments2"].Value;
						attachments3 = config.AppSettings.Settings["attachments3"].Value;
						attachments4 = config.AppSettings.Settings["attachments4"].Value;
						mail_support_error = config.AppSettings.Settings["mail_support_error"].Value;						
					}
					catch (Exception ex)
					{
						MessageBox.Show("Error - " + ex.InnerException+"\n" + ex.StackTrace + "\n" + ex.Message);
					}
				}
				else if(File.Exists(Environment.CurrentDirectory + file[1] + ".config"))
				{
					try
					{

					}
					catch (Exception ex)
					{
						MessageBox.Show("Error - " + ex.InnerException + "\n" + ex.StackTrace + "\n" + ex.Message);
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
				if (Sel_DataTime)
				{
					OdbcDataAdapter adapter1 = new OdbcDataAdapter(Sel_Request_DataTime, cnS11);					
					DataTable table1 = new DataTable();
					adapter1.Fill(table1);
					dataGridView2.DataSource = table1;
					xls = dataGridView2[0, 0].Value.ToString();
				}
				do
				{
					if (Sel_2)
					{
						if (Do == 2)
							Sel_2_request = Sel_3_request;
						else if (Do == 3)
							Sel_2_request = Sel_4_request;
						x = 1;
						Invoke(new Action(Label));
						OdbcDataAdapter adapter = new OdbcDataAdapter(Sel_2_request, cnS11);
						DataTable table = new DataTable();
						adapter.Fill(table);
						dataGridView1.DataSource = table;
						Invoke(new Action(Label));
						//progressBar1.Maximum = dataGridView1.RowCount + dataGridView1.ColumnCount;
						backgroundWorker1.ReportProgress(dataGridView1.RowCount + dataGridView1.ColumnCount);
						backgroundWorker1.ReportProgress(0);
					}
					if (excel_export)
					{
						try
						{
							ExeConfigurationFileMap configFile_excel = new ExeConfigurationFileMap();
							configFile_excel.ExeConfigFilename = Path.Combine(conf1);
							Configuration config_excel = ConfigurationManager.OpenMappedExeConfiguration(configFile_excel, ConfigurationUserLevel.None);
							string row_1 = config_excel.AppSettings.Settings["row_1"].Value;
							var Cells_xlRight = config_excel.AppSettings.Settings["Cells_xlRight"].Value;
							var Cells_xlCenter = config_excel.AppSettings.Settings["Cells_xlCenter"].Value;
							var Cells_xlLeft = config_excel.AppSettings.Settings["Cells_xlLeft"].Value;
							if (Do == 2)							
								conf1 = conf2;							
							if (Do == 3)
								conf1 = conf2;
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
							int f=0;
							var qwert = row_1.Split(',');
							foreach (var row in qwert)
							{
								xlApp.Cells[1,f+ 1] = row;
								
								f++;
							}
							int er=1;
							Excel.Constants xlLeft = Excel.Constants.xlLeft;
							Excel.Constants xlRight = Excel.Constants.xlRight;
							Excel.Constants xlCenter = Excel.Constants.xlCenter;
							var xl = Excel.Constants.xlLeft;							
							foreach (var row in Cells_xlRight.Split(','))
							{
								if (Cells_xlRight != "")
								{
									xl = xlRight;
									er = int.Parse(row);
								}
								xlApp.Cells[1, er].HorizontalAlignment = xl;
							}
							foreach (var row in Cells_xlCenter.Split(','))
							{
								if (Cells_xlCenter != "")
								{
									xl = xlCenter;
									er = int.Parse(row);
								}
								xlApp.Cells[1, er].HorizontalAlignment = xl;
							}
							foreach (var row in Cells_xlLeft.Split(','))
							{
								if (Cells_xlLeft != "")
								{
									xl = xlLeft;
									er = int.Parse(row);
								}
								xlApp.Cells[1, er].HorizontalAlignment = xl;
							}
							//xlApp.Cells[1, 1] = qwer;
							//xlApp.Cells[1, er].HorizontalAlignment = Excel.Constants.xlRight;
							//xlApp.Cells[1, 2] = xls;
							//xlApp.Cells[2, 1] = xlApp_Cells_2_1;
							//xlApp.Cells[2, 2] = xlApp_Cells_2_2;
							//xlApp.Cells[2, 3] = xlApp_Cells_2_3;
							//xlApp.Cells[2, 4] = xlApp_Cells_2_4;
							//xlWorkSheet.Cells[1, 4].EntireRow.Font.Bold = xlWorkSheet.Cells[2, 4].EntireRow.Font.Bold = xlApp_Cells_1_4_2_4_Font;
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
							xlWorkBook.SaveAs(attachments1, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
							xlWorkBook.Close(true, misValue, misValue);
							xlApp.Quit();

							ReleaseObject(xlWorkSheet);
							ReleaseObject(xlWorkBook);
							ReleaseObject(xlApp);
							
						}
						catch(Exception ex)
						{

							MessageBox.Show(conf1 + ex.InnerException + "\n" + ex.StackTrace + "\n" + ex.Message);
						}
					}
					if (Do == attachments)
						break;
					Do++;
				}
				while (attachments > 1);
				try
				{
					if (mail_export)
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
						catch (SmtpException ex)
						{
							x = 7;
							Invoke(new Action(Label));
							backgroundWorker1.CancelAsync();
							notifyIcon1.BalloonTipIcon = ToolTipIcon.Error;
							notifyIcon1.BalloonTipText = "SMTP Error ";
							notifyIcon1.ShowBalloonTip(5000);
							//label1.Text = "Ошибка.";
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
						finally
						{
							Message.Attachments.Dispose();
							x = 8;
							Invoke(new Action(Label));
							x = 9;
							Invoke(new Action(Label));
							if (attachments == 1)
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
						MessageBox.Show("Ошибка!", "smtp" );
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
