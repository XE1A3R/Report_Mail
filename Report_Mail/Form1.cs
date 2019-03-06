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
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Report_Mail
{
	public partial class Form1 : Form
	{
		private int x = 0;
		private string temp;
		private bool Sel_DataTime;
		private string Sel_Request_DataTime;
		private bool Sel_2;
		private string Sel_2_request;
		private string Sel_3_request;
		private string Sel_4_request;
		private string Sel_5_request;
		private bool excel_export;
		private string conf1;
		private string conf2;
		private string conf3;
		private string conf4;
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
		private string configuration;
		static string ok = "Ok";
		static string error = "Ok, Присутствуют ошибки";
		static string error_1 = "There are errors";
		string MesSub = ok;
		string mysql = error_1;
		string excel = error_1;
		string mail = error_1;
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
						configuration = file[1];
						ExeConfigurationFileMap configFile = new ExeConfigurationFileMap
						{
							ExeConfigFilename = Path.Combine(@"C:\confs\", file[1] + ".config")
						};
						Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFile, ConfigurationUserLevel.None);
						temp = config.AppSettings.Settings["temp"].Value;
						Sel_DataTime = bool.Parse(config.AppSettings.Settings["Sel_DataTime"].Value);
						Sel_Request_DataTime = config.AppSettings.Settings["Sel_Request_DataTime"].Value;
						Sel_2 = bool.Parse(config.AppSettings.Settings["sel_2"].Value);
						Sel_2_request = config.AppSettings.Settings["Sel_2_request"].Value;
						Sel_3_request = config.AppSettings.Settings["Sel_3_request"].Value;
						Sel_4_request = config.AppSettings.Settings["Sel_4_request"].Value;
						Sel_5_request = config.AppSettings.Settings["Sel_5_request"].Value;
						excel_export = bool.Parse(config.AppSettings.Settings["excel_export"].Value);
						conf1 = config.AppSettings.Settings["conf1"].Value;
						conf2 = config.AppSettings.Settings["conf2"].Value;
						conf3 = config.AppSettings.Settings["conf3"].Value;
						conf4 = config.AppSettings.Settings["conf4"].Value;
						mail_export = bool.Parse(config.AppSettings.Settings["mail_export"].Value);
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
						var today = DateTime.Today;
						var day_old = Convert.ToInt32(today.DayOfWeek) + 6;
						var monday_old = today.AddDays(-day_old);
						var sunday_old = monday_old.AddDays(6);
						Log.logger.Error("Error User:{0}, ID:{1}, IP:{2}, Ver:{3} \n" +
						"" + Environment.NewLine + "" +
						"" + Environment.NewLine + "" +
						" {4}" +
						"" + Environment.NewLine + "" +
						" {5}" +
						"" + Environment.NewLine + "" +
						"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, ex.Message, ex.StackTrace);

						SmtpClient smtp = new SmtpClient("mail.gb15.ru", 25)
						{
							Credentials = new NetworkCredential("robot@gb15.ru", "1oc@1RoBoT")
						};
						MailMessage Message = new MailMessage
						{
							From = new MailAddress("robot@gb15.ru", "Robot")
						};
						Message.To.Add(new MailAddress("mis@gb15.ru"));
						Message.Subject = "Report_Mail - Error " + configuration;
						Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
						" " + ex.Message + "";
						Message.Attachments.Add(new Attachment("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"));
						try
						{
							smtp.Send(Message);
						}
						catch (SmtpException)
						{
							MessageBox.Show("Ошибка!", "smtp");
						}
						MessageBox.Show("Error - " + ex.InnerException + "\n" + ex.StackTrace + "\n" + ex.Message);
					}
				}
				else if (File.Exists(Environment.CurrentDirectory + file[1] + ".config"))
				{
					try
					{
						var configuration = file[1];
						ExeConfigurationFileMap configFile = new ExeConfigurationFileMap
						{
							ExeConfigFilename = Path.Combine(@"C:\confs\", file[1] + ".config")
						};
						Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFile, ConfigurationUserLevel.None);
						temp = config.AppSettings.Settings["temp"].Value;
						Sel_DataTime = bool.Parse(config.AppSettings.Settings["Sel_DataTime"].Value);
						Sel_Request_DataTime = config.AppSettings.Settings["Sel_Request_DataTime"].Value;
						Sel_2 = bool.Parse(config.AppSettings.Settings["sel_2"].Value);
						Sel_2_request = config.AppSettings.Settings["Sel_2_request"].Value;
						Sel_3_request = config.AppSettings.Settings["Sel_3_request"].Value;
						Sel_4_request = config.AppSettings.Settings["Sel_4_request"].Value;
						Sel_5_request = config.AppSettings.Settings["Sel_5_request"].Value;
						excel_export = bool.Parse(config.AppSettings.Settings["excel_export"].Value);
						conf1 = config.AppSettings.Settings["conf1"].Value;
						conf2 = config.AppSettings.Settings["conf2"].Value;
						conf3 = config.AppSettings.Settings["conf3"].Value;
						conf4 = config.AppSettings.Settings["conf4"].Value;
						mail_export = bool.Parse(config.AppSettings.Settings["mail_export"].Value);
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
						var today = DateTime.Today;
						var day_old = Convert.ToInt32(today.DayOfWeek) + 6;
						var monday_old = today.AddDays(-day_old);
						var sunday_old = monday_old.AddDays(6);
						Log.logger.Error("Error User:{0}, ID:{1}, IP:{2}, Ver:{3} \n" +
						"" + Environment.NewLine + "" +
						"" + Environment.NewLine + "" +
						" {4}" +
						"" + Environment.NewLine + "" +
						" {5}" +
						"" + Environment.NewLine + "" +
						"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, ex.Message, ex.StackTrace);

						SmtpClient smtp = new SmtpClient("mail.gb15.ru", 25)
						{
							Credentials = new NetworkCredential("robot@gb15.ru", "1oc@1RoBoT")
						};
						MailMessage Message = new MailMessage
						{
							From = new MailAddress("robot@gb15.ru", "Robot")
						};
						Message.To.Add(new MailAddress("mis@gb15.ru"));
						Message.Subject = "Report_Mail - Error " + configuration;
						Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
						" " + ex.Message + "";
						Message.Attachments.Add(new Attachment("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"));
						try
						{
							smtp.Send(Message);
						}
						catch (SmtpException)
						{
							MessageBox.Show("Ошибка!", "smtp");
						}
						MessageBox.Show("Error - " + ex.InnerException + "\n" + ex.StackTrace + "\n" + ex.Message);
					}
				}
				else
				{
					Exception ex = new Exception("Файл " + file[1] + ".config ненайден");
					var today = DateTime.Today;
					var day_old = Convert.ToInt32(today.DayOfWeek) + 6;
					var monday_old = today.AddDays(-day_old);
					var sunday_old = monday_old.AddDays(6);
					Log.logger.Error("Error User:{0}, ID:{1}, IP:{2}, Ver:{3} \n" +
					"" + Environment.NewLine + "" +
					"" + Environment.NewLine + "" +
					" {4}" +
					"" + Environment.NewLine + "" +
					" {5}" +
					"" + Environment.NewLine + "" +
					"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, ex.Message, ex.StackTrace);

					SmtpClient smtp = new SmtpClient("mail.gb15.ru", 25)
					{
						Credentials = new NetworkCredential("robot@gb15.ru", "1oc@1RoBoT")
					};
					MailMessage Message = new MailMessage
					{
						From = new MailAddress("robot@gb15.ru", "Robot")
					};
					Message.To.Add(new MailAddress("mis@gb15.ru"));
					Message.Subject = "Report_Mail - Error " + configuration;
					Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
					" " + ex.Message + "";
					Message.Attachments.Add(new Attachment("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"));
					try
					{
						smtp.Send(Message);
					}
					catch (SmtpException)
					{
						MessageBox.Show("Ошибка!", "smtp");
					}
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
				label1.Text = "Выгрузка в EXCEL... ";
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
			if (Sel_DataTime)
			{
				try
				{
					OdbcDataAdapter adapter1 = new OdbcDataAdapter(Sel_Request_DataTime, cnS11);
					DataTable table1 = new DataTable();
					adapter1.Fill(table1);
					dataGridView2.DataSource = table1;
					xls = dataGridView2[0, 0].Value.ToString();
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
					" {4}" +
					"" + Environment.NewLine + "" +
					" {5}" +
					"" + Environment.NewLine + "" +
					"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, ex.Message, ex.StackTrace);

					SmtpClient smtp = new SmtpClient(smtpClient, smtpClient_port)
					{
						Credentials = new NetworkCredential(from_mail, from_Password)
					};
					MailMessage Message = new MailMessage
					{
						From = new MailAddress(from_mail, from_mail_name)
					};
					Message.To.Add(new MailAddress(mail_support_error));
					Message.Subject = "Report_Mail - Error " + configuration;
					Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
					" " + ex.Message + "";
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
			do
			{
				if (Sel_2)
				{
					try
					{
						if (Do == 2)
							Sel_2_request = Sel_3_request;
						else if (Do == 3)
							Sel_2_request = Sel_4_request;
						else if (Do == 4)
							Sel_2_request = Sel_5_request;

						x = 1;
						Invoke(new Action(Label));
						OdbcDataAdapter adapter = new OdbcDataAdapter(Sel_2_request, cnS11);
						DataTable table = new DataTable();
						adapter.Fill(table);
						dataGridView1.DataSource = table;
						Invoke(new Action(Label));
						//progressBar1.Maximum = dataGridView1.RowCount + dataGridView1.ColumnCount;
						backgroundWorker1.ReportProgress(dataGridView1.RowCount);
						backgroundWorker1.ReportProgress(0);
						mysql = ok;
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
						" {4}" +
						"" + Environment.NewLine + "" +
						" {5}" +
						"" + Environment.NewLine + "" +
						"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, ex.Message, ex.StackTrace);

						SmtpClient smtp = new SmtpClient(smtpClient, smtpClient_port)
						{
							Credentials = new NetworkCredential(from_mail, from_Password)
						};
						MailMessage Message = new MailMessage
						{
							From = new MailAddress(from_mail, from_mail_name)
						};
						Message.To.Add(new MailAddress(mail_support_error));
						Message.Subject = "Report_Mail - Error " + configuration;
						Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
						" " + ex.Message + "";
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
				if (excel_export)
				{
					try
					{
						var att_up = "";
						if (Do == 1)
							att_up = attachments1;
						else if (Do == 2)
						{
							conf1 = conf2;
							att_up = attachments2;
						}
						else if (Do == 3)
						{
							conf1 = conf3;
							att_up = attachments3;
						}
						else if (Do == 4)
						{
							conf1 = conf4;
							att_up = attachments4;
						}
						ExeConfigurationFileMap configFile_excel = new ExeConfigurationFileMap
						{
							ExeConfigFilename = Path.Combine(conf1)
						};
						Configuration config_excel = ConfigurationManager.OpenMappedExeConfiguration(configFile_excel, ConfigurationUserLevel.None);
						string row_1 = config_excel.AppSettings.Settings["row_1"].Value;
						var Cells_xlRight_1 = config_excel.AppSettings.Settings["Cells_xlRight_1"].Value;
						var Cells_xlCenter_1 = config_excel.AppSettings.Settings["Cells_xlCenter_1"].Value;
						var Cells_xlLeft_1 = config_excel.AppSettings.Settings["Cells_xlLeft_1"].Value;
						var Cells_xlRight_2 = config_excel.AppSettings.Settings["Cells_xlRight_2"].Value;
						var Cells_xlCenter_2 = config_excel.AppSettings.Settings["Cells_xlCenter_2"].Value;
						var Cells_xlLeft_2 = config_excel.AppSettings.Settings["Cells_xlLeft_2"].Value;
						var int_h = int.Parse(config_excel.AppSettings.Settings["h"].Value);
						var color = config_excel.AppSettings.Settings["color"].Value;
						var value_color = config_excel.AppSettings.Settings["value_color"].Value;
						var oper = config_excel.AppSettings.Settings["oper"].Value;
						var Red = int.Parse(config_excel.AppSettings.Settings["Red"].Value);
						var Green = int.Parse(config_excel.AppSettings.Settings["Green"].Value);
						var Blue = int.Parse(config_excel.AppSettings.Settings["Blue"].Value);
						var Red_1 = int.Parse(config_excel.AppSettings.Settings["Red_1"].Value);
						var Green_1 = int.Parse(config_excel.AppSettings.Settings["Green_1"].Value);
						var Blue_1 = int.Parse(config_excel.AppSettings.Settings["Blue_1"].Value);
						var data_1 = config_excel.AppSettings.Settings["data_1"].Value;
						var data_2 = config_excel.AppSettings.Settings["data_2"].Value;

						//Excel.Application xlApp;
						//Excel.Workbook xlWorkBook;
						//Excel.Worksheet xlWorkSheet;
						object misValue = System.Reflection.Missing.Value;
						x = 2;
						Invoke(new Action(Label));
						//label1.Text = "Создание нового файла EXCEL...";
						Int16 i, j;
						int h = int_h;
						//xlApp = new Excel.Application();
						//xlWorkBook = xlApp.Workbooks.Add(misValue);
						//xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
						string path = "C:\\TestFileSave\\ABC.xlsx";
						FileInfo filePath = new FileInfo(path);
						var package = new ExcelPackage();

						ExcelWorksheet xlWorkSheet = package.Workbook.Worksheets.Add("Отчет");


						int f = 0;
						h -= 1;
						foreach (var row in row_1.Split(','))
						{
							xlWorkSheet.Cells[h, f + 1].Value = row;
							xlWorkSheet.Cells[h, f + 1].Style.Font.Bold = true;
							f++;
						}
						if (data_1 != "" | data_2 != "")
							xlWorkSheet.Cells[int.Parse(data_1), int.Parse(data_2)].Value = xls;
						var xlLeft = ExcelHorizontalAlignment.Left;
						var xlRight = ExcelHorizontalAlignment.Right;
						var xlCenter = ExcelHorizontalAlignment.Center;
						var xl = ExcelHorizontalAlignment.Left;
						foreach (var row in Cells_xlRight_1.Split(','))
						{
							if (Cells_xlRight_1 != "")
							{
								xl = xlRight;
								h = int.Parse(row);

								xlWorkSheet.Cells[1, h].Style.HorizontalAlignment = xl;
							}
						}
						foreach (var row in Cells_xlCenter_1.Split(','))
						{
							if (Cells_xlCenter_1 != "")
							{
								xl = xlCenter;
								h = int.Parse(row);

								xlWorkSheet.Cells[1, h].Style.HorizontalAlignment = xl;
							}
						}
						foreach (var row in Cells_xlLeft_1.Split(','))
						{
							if (Cells_xlLeft_1 != "")
							{
								xl = xlLeft;
								h = int.Parse(row);

								xlWorkSheet.Cells[1, h].Style.HorizontalAlignment = xl;
							}
						}
						foreach (var row in Cells_xlRight_2.Split(','))
						{
							if (Cells_xlRight_2 != "")
							{
								xl = xlRight;
								h = int.Parse(row);

								xlWorkSheet.Cells[2, h].Style.HorizontalAlignment = xl;
							}
						}
						foreach (var row in Cells_xlCenter_2.Split(','))
						{
							if (Cells_xlCenter_2 != "")
							{
								xl = xlCenter;
								h = int.Parse(row);

								xlWorkSheet.Cells[2, h].Style.HorizontalAlignment = xl;
							}
						}
						foreach (var row in Cells_xlLeft_2.Split(','))
						{
							if (Cells_xlLeft_2 != "")
							{
								xl = xlLeft;
								h = int.Parse(row);

								xlWorkSheet.Cells[2, h].Style.HorizontalAlignment = xl;
							}
						}
						int color_cl = 10;
						foreach (var row in color.Split(','))
						{
							if (color != "")
							{
								color_cl = int.Parse(row);
							}
						}
						x = 3;
						Invoke(new Action(Label));
						for (f = 0; f < this.dataGridView1.Columns.Count; f++)
						{
							xlWorkSheet.Cells[int_h, f + 1].Value = this.dataGridView1.Columns[f].HeaderCell.Value.ToString();
							xlWorkSheet.Cells[int_h, f + 1].Style.Font.Bold = true;
							xlWorkSheet.Cells[int_h, f + 1].Style.Border.Top.Style = ExcelBorderStyle.Medium;
							xlWorkSheet.Cells[int_h, f + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
							xlWorkSheet.Cells[int_h, f + 1].Style.Border.Left.Style = ExcelBorderStyle.Medium;
							xlWorkSheet.Cells[int_h, f + 1].Style.Border.Right.Style = ExcelBorderStyle.Medium;
							
							//xlWorkSheet.Cells[int_h, f + 1].Borders.Weight = Excel.XlBorderWeight.xlMedium;
						}
						//label1.Text = "Выгрузка в EXCEL...";						
						//xlApp.Visible = true;
						xlWorkSheet.View.FreezePanes(int_h+1, 1);
						for (i = 0; i < dataGridView1.RowCount; i++)
						{
							for (j = 0; j < dataGridView1.ColumnCount; j++)
							{
								backgroundWorker1.ReportProgress(int_h);
								xlWorkSheet.Cells[int_h + 1, j + 1].Value = dataGridView1[j, i].FormattedValue.ToString();
								xlWorkSheet.Cells[int_h + 1, j + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
								xlWorkSheet.Cells[int_h + 1, j + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
								xlWorkSheet.Cells[int_h + 1, j + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
								xlWorkSheet.Cells[int_h + 1, j + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
								if (color != "")
								{
									//int test = Convert.ToInt32(value_color);
									//int qwer = Convert.ToInt32(xlWorkSheet.Cells[h + 1, color_cl].Text);
									int cell;
									int.TryParse(xlWorkSheet.Cells[int_h + 1, color_cl].Text, out cell);
									if (oper == "=")
									{
										if (xlWorkSheet.Cells[int_h + 1, color_cl].Text == value_color)
											xlWorkSheet.Row(int_h + 1).Style.Font.Color.SetColor(1, Red, Green, Blue);
										else
											xlWorkSheet.Cells[int_h + 1, color_cl].Style.Font.Color.SetColor(1, Red_1, Green_1, Blue_1);
									}
									else if (oper == ">")
									{
										if (cell > int.Parse(value_color))
											xlWorkSheet.Row(int_h + 1).Style.Font.Color.SetColor(1, Red, Green, Blue);
										else
											xlWorkSheet.Cells[int_h + 1, color_cl].Style.Font.Color.SetColor(1, Red_1, Green_1, Blue_1);
									}
									else if (oper == "<")
									{
										if (cell < int.Parse(value_color))
											xlWorkSheet.Row(int_h + 1).Style.Font.Color.SetColor(1, Red, Green, Blue);
										else
											xlWorkSheet.Cells[int_h + 1, color_cl].Style.Font.Color.SetColor(1, Red_1, Green_1, Blue_1);
									}
									else if (oper == "!=")
									{
										if (cell != int.Parse(value_color))
											xlWorkSheet.Row(int_h + 1).Style.Font.Color.SetColor(1, Red, Green, Blue);
										else
											xlWorkSheet.Cells[int_h + 1, color_cl].Style.Font.Color.SetColor(1, Red_1, Green_1, Blue_1);
									}

								}
								// xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
								//xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные
								//xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
								//xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
								//xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
								//xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
							}
							int_h++;
						}
						//backgroundWorker1.ReportProgress(dataGridView1.ColumnCount + dataGridView1.RowCount);

						for (int t = 1; t < 20; t++)
						{
							xlWorkSheet.Cells.AutoFitColumns();
							//((Excel.Range)xlWorkSheet.Columns[t]).AutoFit();
						}
						x = 4;
						Invoke(new Action(Label));
						//label1.Text = "Сохранение EXCEL...";
						Directory.CreateDirectory(@temp);
						FileStream aFile = new FileStream(att_up, FileMode.Create);
						package.SaveAs(aFile);
						package.Dispose();
						aFile.Close();
						//xlWorkBook.SaveAs(att_up, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

						//xlWorkBook.Close(true, misValue, misValue);
						//xlApp.Quit();

						//ReleaseObject(xlWorkSheet);
						//ReleaseObject(xlWorkBook);
						//ReleaseObject(xlApp);
						excel = ok;
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
						" {4}" +
						"" + Environment.NewLine + "" +
						" {5}" +
						"" + Environment.NewLine + "" +
						"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, ex.Message, ex.StackTrace);

						SmtpClient smtp = new SmtpClient(smtpClient, smtpClient_port)
						{
							Credentials = new NetworkCredential(from_mail, from_Password)
						};
						MailMessage Message = new MailMessage
						{
							From = new MailAddress(from_mail, from_mail_name)
						};
						Message.To.Add(new MailAddress(mail_support_error));
						Message.Subject = "Report_Mail - Error " + configuration;
						Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
						" " + ex.Message + "";
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
				x = 9;
				Invoke(new Action(Label));
				if (Do == attachments)
					break;
				Do++;
			}
			while (attachments > 1);

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
						Message.Attachments.Add(new Attachment(attachments1));
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
					//Good();
					mail = ok;
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
					" {4}" +
					"" + Environment.NewLine + "" +
					" {5}" +
					"" + Environment.NewLine + "" +
					"" + Environment.NewLine + "", Log.username, Data.Person_id, Log.IP_Address, Log.version, ex.Message, ex.StackTrace);

					Message.To.Add(new MailAddress(mail_support_error));
					Message.Subject = "Report_Mail - Error " + configuration;
					Message.Body = "Error User: " + Log.username + ", ID: " + Data.Person_id + ", IP: " + Log.IP_Address + ", Ver: " + Log.version + " \n" +
					" " + ex.Message + "";
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
					if (attachments == 1)
						File.Delete(attachments1);
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
					if (System.IO.Directory.GetDirectories(temp).Length + System.IO.Directory.GetFiles(temp).Length > 0) { }
					else
						Directory.Delete(temp);
					x = 9;
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
			var today = DateTime.Today;
			var day_old = Convert.ToInt32(today.DayOfWeek) + 6;
			var monday_old = today.AddDays(-day_old);
			var sunday_old = monday_old.AddDays(6);
			SmtpClient smtp1 = new SmtpClient(smtpClient, smtpClient_port)
			{
				Credentials = new NetworkCredential(from_mail, from_Password)
			};
			MailMessage Message1 = new MailMessage
			{
				From = new MailAddress(from_mail, from_mail_name)
			};
			Message1.To.Add(new MailAddress(mail_support_error));
			if (File.Exists("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"))
				MesSub = error;
			Message1.Subject = "Report_Mail - " + MesSub;
			Message1.Body = "Config - " + configuration +
			Environment.NewLine +
			"MySql - " + mysql +
			Environment.NewLine +
			"Excel - " + excel +
			Environment.NewLine +
			"Mail - " + mail +
			Environment.NewLine +
			Environment.NewLine +
			Sel_2_request + Environment.NewLine + Sel_3_request + Environment.NewLine + Sel_4_request + Environment.NewLine + Sel_5_request + Environment.NewLine;
			if (File.Exists("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"))
				Message1.Attachments.Add(new Attachment("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"));
			try
			{
				smtp1.Send(Message1);
			}
			catch (SmtpException)
			{
				MessageBox.Show("Ошибка!", "smtp");
			}
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

		//void Good()
		//{

		// var today = DateTime.Today;
		// var day_old = Convert.ToInt32(today.DayOfWeek) + 6;
		// var monday_old = today.AddDays(-day_old);
		// var sunday_old = monday_old.AddDays(6);
		// SmtpClient smtp1 = new SmtpClient(smtpClient, smtpClient_port)
		// {
		// Credentials = new NetworkCredential(from_mail, from_Password)
		// };
		// MailMessage Message1 = new MailMessage
		// {
		// From = new MailAddress(from_mail, from_mail_name)
		// };
		// Message1.To.Add(new MailAddress(mail_support_error));
		// var ok = "Ok";
		// var error = "Ok, Присутствуют ошибки";
		// var MesSub = ok;
		// var mysql = ok;
		// var excel = ok;
		// var mail = ok;
		// Message1.Body = Sel_2_request + " " + Sel_3_request + " " + Sel_4_request + " " + Sel_5_request;
		// if (File.Exists("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"))
		// MesSub = error;
		// Message1.Subject = "Report_Mail - " + MesSub;
		// Message1.Body = "MySql - " +mysql+
		// Environment.NewLine +
		// "Excel - " + excel+
		// Environment.NewLine +
		// "Mail - " +mail +
		// Environment.NewLine +
		// Environment.NewLine+
		// Sel_2_request + Environment.NewLine + Sel_3_request + Environment.NewLine + Sel_4_request + Environment.NewLine + Sel_5_request + Environment.NewLine;
		// if (File.Exists("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"))
		// Message1.Attachments.Add(new Attachment("" + Environment.CurrentDirectory + "/logs/" + today.ToString("yyyy-MM-dd") + ".log"));
		// try
		// {
		// smtp1.Send(Message1);
		// }
		// catch (SmtpException)
		// {
		// MessageBox.Show("Ошибка!", "smtp");
		//}
		//}
	}
}
