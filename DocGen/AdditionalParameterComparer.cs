using System.Collections.Generic;
using SharePointPnP.PowerShell.CmdletHelpAttributes;

namespace DocGen
{
    internal class AdditionalParameterComparer : IEqualityComparer<CmdletAdditionalParameter>
    {
        public bool Equals(CmdletAdditionalParameter x, CmdletAdditionalParameter y)
        {
            return x.ParameterName.Equals(y.ParameterName);
        }

        public int GetHashCode(CmdletAdditionalParameter obj)
        {
            return obj.ParameterName.GetHashCode();
        }
    }

}