#nullable enable
using System.Collections.Generic;
using System.Text.Json;
using Report_Mail.Interface;
using Report_Mail.Model;

namespace Report_Mail.Controller
{
    public class ConfigAppController : BaseConfigController
    {
        private static string? _configFile;
        public ConfigJson? ConfigJson { get; }

        public ConfigAppController(IReadOnlyList<string> file) : base(file)
        {
            _configFile = CurrentConfig;
            ConfigJson = Deserialize();
        }

        private static ConfigJson? Deserialize()
        {
            if (_configFile == null) return null;
            var jsonString = System.IO.File.ReadAllText(_configFile);
            return JsonSerializer.Deserialize<ConfigJson>(jsonString);
        }
    }
}