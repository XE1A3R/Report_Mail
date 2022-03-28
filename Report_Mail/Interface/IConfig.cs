using System.Collections.Generic;

namespace Report_Mail.Interface
{
    public interface IConfig
    {
        string GetConfig(IReadOnlyList<string> file);
    }
}