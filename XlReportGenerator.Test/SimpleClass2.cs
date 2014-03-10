using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlReportGenerator.Test
{
    public class SimpleClass2
    {
        
        public String Name { get; set; }

        [ColumnName("Age")]
        public Int64 Age { get; set; }

        [ColumnName("Date of Birth")]
        [DateFormat("dd MMM yyyy")]
        public DateTime BOD { get; set; }
    }
}
