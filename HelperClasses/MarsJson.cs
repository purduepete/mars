using System.Collections.Generic;

namespace MarsAddIn.HelperClasses
{
    public class MarsSettings
    {
        public string SavedFileLocation { get; set; }
        public List<string> DatabaseServers { get; set; }
        public Dictionary<string, bool> RibbonSettings { get; set; }
    }

    public class SavedFiles
    {
        public Dictionary<string, string> SavedFilesLocation { get; set; }
    }
}
