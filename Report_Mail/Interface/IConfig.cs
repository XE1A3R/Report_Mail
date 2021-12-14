using System.Collections.Generic;

namespace Report_Mail.Interface
{
    public interface IConfig
    {
        public string CurrentConfig { get; }
        public IReadOnlyList<string> File { get; }

        string GetConfig(IReadOnlyList<string> file);
    }
}