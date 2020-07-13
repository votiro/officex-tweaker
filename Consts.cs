using System.Collections.Generic;

namespace BypassAv
{
    public static class Consts
    {
        public static HashSet<string> PROPERTY_ENTRIES { get; set; } = new HashSet<string>
        {
            "docProps/app.xml",
            "docProps/core.xml"
        };

        public const string TOP_RELATIONS = "_rels/.rels";

        public const string DEFAULT_OUTPUT_SUFFIX = ".out";
    }
}
