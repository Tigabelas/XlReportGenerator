using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlReportGenerator
{
    [AttributeUsage(AttributeTargets.All, Inherited = false, AllowMultiple = true)]
    public sealed class HyperlinkFormat : Attribute
    {
        public Boolean IsHyperlink { get; private set;}

        public HyperlinkFormat(Boolean isHyperlink)
        {
            this.IsHyperlink = isHyperlink;
        }
    }
}
