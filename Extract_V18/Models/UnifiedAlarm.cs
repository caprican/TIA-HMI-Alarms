using System.Collections.Generic;
using System.Diagnostics;

using Siemens.Engineering.HmiUnified;

namespace Extract.Models
{
    [DebuggerDisplay("{Tagname}")]
    public class UnifiedAlarm
    {
        public HmiSoftware Hmi { get; set; }
        public string ClassName { get; set; } = string.Empty;
        public string Tagname { get; set; } = string.Empty;
        public string Origin { get; set; } = string.Empty;
        public Dictionary<string, string> Descriptions { get; set; }
    }
}
