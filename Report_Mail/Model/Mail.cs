using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using Report_Mail.Interface;

namespace Report_Mail.Model
{
    [Serializable]

    public class Mail : IMail
    {
        [JsonPropertyName("MailExport")] public string MailExport { get; set; }
        [JsonPropertyName("SmtpClient")] public string SmtpClient { get; set; }
        [JsonPropertyName("Port")] public int Port { get; set; }
        [JsonPropertyName("From")] public string From { get; set; }
        [JsonPropertyName("Name")] public string Name { get; set; }
        [JsonPropertyName("Password")] public string Password { get; set; }
        [JsonPropertyName("To")] public List<string> To { get; set; }
        [JsonPropertyName("Cc")] public List<string> Cc { get; set; }
        [JsonPropertyName("Subject")] public string Subject { get; set; }
        [JsonPropertyName("Body")] public string Body { get; set; }
        [JsonPropertyName("MailSupportError")] public List<string> MailSupportError { get; set; }
    }
}