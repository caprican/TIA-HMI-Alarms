
using System.Collections.Generic;

namespace Extract.Core
{
    public class InterfaceMember
    {
        public string Name { get; set; }

        public DirectionMember Direction { get; set; }

        public string Type { get; set; }

        //public string DerivedType { get; set; }

        public Dictionary<string, string> Description { get; set; }

        public string DefaultValue { get; set; } = string.Empty;

        public bool Islocked { get; set; } = false;

        public bool HidenInterface { get; set; } = false;
    }
}
