using System;
using System.Text.Json.Serialization;
using Report_Mail.Interface;

namespace Report_Mail.Model
{
    [Serializable]
    public class LocationTable : ILocationTable 
    {
        [JsonPropertyName("PrintHeaders")]
        public bool PrintHeaders { get; set; }
        [JsonPropertyName("SmartTable")]
        public bool SmartTable { get; set; }
        [JsonPropertyName("FreezePanes")]
        public bool FreezePanes { get; set; }
        [JsonPropertyName("FontSize")]
        public uint Size { get; set; }
        [JsonPropertyName("Request")]
        public string Request { get; set; }
        [JsonPropertyName("Column")]
        public int Column { get; set; }
        [JsonPropertyName("Row")]
        public int Row { get; set; }
    }
}