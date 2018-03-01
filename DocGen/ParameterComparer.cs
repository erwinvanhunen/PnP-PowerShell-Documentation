using DocGen.Model;
using System.Collections.Generic;

namespace DocGen
{
    internal class ParameterComparer : IEqualityComparer<CmdletParameterInfo>
    {
        public bool Equals(CmdletParameterInfo x, CmdletParameterInfo y)
        {
            return x.Name.Equals(y.Name);
        }

        public int GetHashCode(CmdletParameterInfo obj)
        {
            return obj.Name.GetHashCode();
        }
    }

}