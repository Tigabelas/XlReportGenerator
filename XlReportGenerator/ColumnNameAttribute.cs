using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlReportGenerator
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = true)]
    public sealed class ColumnName : Attribute
    {
        public String Name
        {
            get; private set;
        }

        // This is a positional argument
        public ColumnName(String columnName)
        {
            this.Name = columnName;
        }
    }
}
