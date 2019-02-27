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
                        backgroundWorker1.ReportProgress(dataGridView1.RowCount + dataGridView1.ColumnCount);
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
                        string row_2 = config_excel.AppSettings.Settings["row_2"].Value;
                        var Cells_xlRight_1 = config_excel.AppSettings.Settings["Cells_xlRight_1"].Value;
                        var Cells_xlCenter_1 = config_excel.AppSettings.Settings["Cells_xlCenter_1"].Value;
                        var Cells_xlLeft_1 = config_excel.AppSettings.Settings["Cells_xlLeft_1"].Value;
                        var Cells_xlRight_2 = config_excel.AppSettings.Settings["Cells_xlRight_2"].Value;
                        var Cells_xlCenter_2 = config_excel.AppSettings.Settings["Cells_xlCenter_2"].Value;
                        var Cells_xlLeft_2 = config_excel.AppSettings.Settings["Cells_xlLeft_2"].Value;
                        var int_h = int.Parse(config_excel.AppSettings.Settings["h"].Value);
                        var color = config_excel.AppSettings.Settings["color"].Value;
                        var data_1 = config_excel.AppSettings.Settings["data_1"].Value;
                        var data_2 = config_excel.AppSettings.Settings["data_2"].Value;

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
                        int f = 0;
                        foreach (var row in row_1.Split(','))
                        {
                            xlApp.Cells[1, f + 1] = row;
                            xlApp.Cells[1, f + 1].EntireRow.Font.Bold = true;
                            f++;
                        }
                        f = 0;
                        foreach (var row in row_2.Split(','))
                        {
                            xlApp.Cells[2, f + 1] = row;
                            xlApp.Cells[2, f + 1].EntireRow.Font.Bold = true;
                            xlWorkSheet.Cells[2, f + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // все стороны
                            xlWorkSheet.Cells[2, f + 1].Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            //xlWorkSheet.Cells[2, f + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
                            //xlWorkSheet.Cells[2, f + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные
                            //xlWorkSheet.Cells[2, f + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                            //xlWorkSheet.Cells[2, f + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                            //xlWorkSheet.Cells[2, f + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                            //xlWorkSheet.Cells[2, f + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            f++;
                        }
                        int er = 0;

                        if (data_1 != "" | data_2 != "")
                            xlApp.Cells[int.Parse(data_1), int.Parse(data_2)] = xls;
                        Excel.Constants xlLeft = Excel.Constants.xlLeft;
                        Excel.Constants xlRight = Excel.Constants.xlRight;
                        Excel.Constants xlCenter = Excel.Constants.xlCenter;
                        var xl = Excel.Constants.xlLeft;
                        foreach (var row in Cells_xlRight_1.Split(','))
                        {
                            if (Cells_xlRight_1 != "")
                            {
                                xl = xlRight;
                                er = int.Parse(row);

                                xlApp.Cells[1, er].HorizontalAlignment = xl;
                            }
                        }
                        foreach (var row in Cells_xlCenter_1.Split(','))
                        {
                            if (Cells_xlCenter_1 != "")
                            {
                                xl = xlCenter;
                                er = int.Parse(row);

                                xlApp.Cells[1, er].HorizontalAlignment = xl;
                            }
                        }
                        foreach (var row in Cells_xlLeft_1.Split(','))
                        {
                            if (Cells_xlLeft_1 != "")
                            {
                                xl = xlLeft;
                                er = int.Parse(row);

                                xlApp.Cells[1, er].HorizontalAlignment = xl;
                            }
                        }
                        foreach (var row in Cells_xlRight_2.Split(','))
                        {
                            if (Cells_xlRight_2 != "")
                            {
                                xl = xlRight;
                                er = int.Parse(row);

                                xlApp.Cells[2, er].HorizontalAlignment = xl;
                            }
                        }
                        foreach (var row in Cells_xlCenter_2.Split(','))
                        {
                            if (Cells_xlCenter_2 != "")
                            {
                                xl = xlCenter;
                                er = int.Parse(row);

                                xlApp.Cells[2, er].HorizontalAlignment = xl;
                            }
                        }
                        foreach (var row in Cells_xlLeft_2.Split(','))
                        {
                            if (Cells_xlLeft_2 != "")
                            {
                                xl = xlLeft;
                                er = int.Parse(row);

                                xlApp.Cells[2, er].HorizontalAlignment = xl;
                            }
                        }
                        int color_cl = 0;
                        foreach (var row in color.Split(','))
                        {
                            if (color != "")
                            {
                                color_cl = int.Parse(row);
                            }
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

                            for (j = 0; j < dataGridView1.ColumnCount; j++)
                            {
                                backgroundWorker1.ReportProgress(i + j);
                                xlWorkSheet.Cells[h + 1, j + 1] = dataGridView1[j, i].Value.ToString();
                                xlWorkSheet.Cells[h + 1, j + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // все стороны
                                if (color != "")
                                {
                                    if (xlWorkSheet.Cells[h + 1, color_cl].Text == "0")
                                        xlApp.Rows[h + 1].Font.Color = Color.FromArgb(0, 128, 0);
                                    else
                                        xlApp.Cells[h + 1, color_cl].Font.Color = Color.Red;
                                }
                                // xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
                                //xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные
                                //xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                                //xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                                //xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                                //xlWorkSheet.Cells[h + 1, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            }
                            h++;
                        }
                        //backgroundWorker1.ReportProgress(dataGridView1.ColumnCount + dataGridView1.RowCount);

                        for (int t = 1; t < 20; t++)
                        {
                            ((Excel.Range)xlWorkSheet.Columns[t]).AutoFit();
                        }
                        x = 4;
                        Invoke(new Action(Label));
                        //label1.Text = "Сохранение EXCEL...";
                        Directory.CreateDirectory(@temp);
                        xlWorkBook.SaveAs(att_up, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                        xlWorkBook.Close(true, misValue, misValue);
                        xlApp.Quit();

                        ReleaseObject(xlWorkSheet);
                        ReleaseObject(xlWorkBook);
                        ReleaseObject(xlApp);
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
