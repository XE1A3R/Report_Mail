using System;
using System.Collections.Generic;

namespace Report_Mail.Interface
{
    public interface IMail
    {
        public string MailExport { get; set; }
        public string SmtpClient { get; set; }
        public int Port { get; set; }
        public string From { get; set; }
        public string Name { get; set; }
        public string Password { get; set; }
        public List<string> To { get; set; }
        public List<string> Cc { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public List<string> MailSupportError { get; set; }
    }
}