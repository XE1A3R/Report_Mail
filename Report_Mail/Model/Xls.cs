using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using Report_Mail.Interface;

namespace Report_Mail.Model
{
    [Serializable]
    public class Xls : IXls
    {
        [JsonPropertyName("Name")]
        public string Name { get; set; }
        [JsonPropertyName("Sheet")]
        public List<Sheet> Sheets { get; set; }
        [JsonPropertyName("Attachments")]
        public string Attachments { get; set; }
        [JsonPropertyName("Format")]
        public string Format { get; set; }
    }
}
