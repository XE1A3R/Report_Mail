using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using Report_Mail.Interface;

namespace Report_Mail
{
    [Serializable]
    public class Sheet : ISheet
    {
        [JsonPropertyName("Name")]
        public string Name { get; set; }
        [JsonPropertyName("LocationTable")]
        public List<LocationTable> Locations { get; set; }
    }
}