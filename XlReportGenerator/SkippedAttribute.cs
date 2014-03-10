using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlReportGenerator
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = true)]
    public sealed class Skipped : Attribute
    {
        public Boolean IsSkipped { get; set; }

        public Skipped()
        {
            IsSkipped = true;
        }
    }
}
