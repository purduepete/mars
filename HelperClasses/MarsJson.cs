using System.Collections.Generic;

namespace MarsAddIn.HelperClasses
{
    public class MarsSettings
    {
        public List<string> DatabaseServers { get; set; }
        public Dictionary<string, bool> RibbonSettings { get; set; }
    }
}

