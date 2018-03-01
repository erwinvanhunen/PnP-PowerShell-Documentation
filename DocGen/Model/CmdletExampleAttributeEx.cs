using SharePointPnP.PowerShell.CmdletHelpAttributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model
{
    public class CmdletExampleAttributeEx
    {
        public string Code { get; set; }
        /// <summary>
        /// Any introduction text
        /// </summary>
        public string Introduction { get; set; }
        /// <summary>
        /// Any remarks, to be rendered underneath the example
        /// </summary>
        public string Remarks { get; set; }
        /// <summary>
        /// The sort order of the example within the list of all examples.
        /// </summary>
        public int SortOrder { get; set; }

        public List<string> Platforms { get; set; }
    }
}
