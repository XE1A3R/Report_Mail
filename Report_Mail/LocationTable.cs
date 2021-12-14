using System;
using System.Text.Json.Serialization;
using Report_Mail.Interface;

namespace Report_Mail
{
    [Serializable]
    public class LocationTable : ILocationTable 
    {
        [JsonPropertyName("Request")]
        public string Request { get; set; }
        [JsonPropertyName("Column")]
        public int Column { get; set; }
        [JsonPropertyName("Row")]
        public int Row { get; set; }
    }
}