#nullable enable
using System.Collections.Generic;
using System.Text.Json;

namespace Report_Mail
{
    public class ConfigApp : BaseConfig
    {
        private static string? _configFile;
        public ConfigJson? ConfigJson { get; }

        public ConfigApp(IReadOnlyList<string> file) : base(file)
        {
            _configFile = CurrentConfig;
            ConfigJson = Deserialize();
        }

        private static ConfigJson? Deserialize()
        {
            if (_configFile != null)
            {
                var jsonString = System.IO.File.ReadAllText(_configFile);
                return JsonSerializer.Deserialize<ConfigJson>(jsonString);
            }

            return null;
        }
    }
}