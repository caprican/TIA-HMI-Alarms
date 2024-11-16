using System.Diagnostics;

using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HmiUnified.HmiConnections;

namespace Extract.Models
{
    [DebuggerDisplay("{PlcTag}")]
    public class UnifiedTag
    {
        public HmiSoftware Hmi { get; set; }
        public HmiConnection Connexion { get; set; }
        public string PlcTag { get; set; } = string.Empty;
        public string Tagname { get; set; } = string.Empty;
        public string Folder { get; set; }
    }
}
