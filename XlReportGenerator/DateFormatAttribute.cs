using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlReportGenerator
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = true)]
    public sealed class DateFormat : Attribute
    {
        public String Format { get; private set;}

        public DateFormat(String format)
        {
            this.Format = format;
        }
    }
}
