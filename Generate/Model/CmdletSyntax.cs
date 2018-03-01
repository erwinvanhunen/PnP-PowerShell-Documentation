using System.Collections.Generic;

namespace SharePointPnP.PowerShell.ModuleFilesGenerator.Model
{
    public class CmdletSyntax
    {
        public string ParameterSetName { get; set; }
        public List<CmdletParameterInfo> Parameters { get; set; }

        public CmdletSyntax()
        {
            Parameters = new List<CmdletParameterInfo>();
        }

        public override int GetHashCode()
        {
            return (string.Format("{0}|{1}",
               ParameterSetName,
               Parameters.GetHashCode()
           ).GetHashCode());
        }
    }
}