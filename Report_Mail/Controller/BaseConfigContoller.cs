using System;
using System.Collections.Generic;
using System.IO;
using Report_Mail.Interface;

namespace Report_Mail.Controller
{
    public class BaseConfigController : IConfig
    {
        protected string CurrentConfig { get; }
        private IReadOnlyList<string> File { get; }

        protected BaseConfigController(IReadOnlyList<string> file)
        {
	        File = file ?? throw new ArgumentNullException(nameof(file));
	        CurrentConfig = GetConfig(file);
        }
        
        public string GetConfig(IReadOnlyList<string> file)
        {
            return System.IO.File.Exists($@"{Directory.GetCurrentDirectory()}\{file[0]}.json") ? $@"{Directory.GetCurrentDirectory()}\{file[0]}.json" : null;
        }
    }
}