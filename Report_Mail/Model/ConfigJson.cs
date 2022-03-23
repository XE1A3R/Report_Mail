using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using Report_Mail.Controller;
using Report_Mail.Interface;

namespace Report_Mail.Model
{
    [Serializable]
    public class ConfigJson : IConfigJson
    {
        [JsonPropertyName("Xls")] 
        public List<Xls> Xls { get; set; }
        [JsonPropertyName("Mail")]
        public List<Mail> Mail { get; set; }
    }
}
