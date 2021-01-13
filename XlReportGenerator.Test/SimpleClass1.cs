using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlReportGenerator.Test
{
    public class SimpleClass1
    {
        [ColumnName("Kolom 1")]
        public String Field1 { get; set; }

        [ColumnName("Kolom 2")]
        public String Field2 { get; set; }

        [Skipped("Sheet1, Sheet2")]
        public String Field3 { get; set; }

        [ColumnName("Kolom 4")]
        public Decimal Field4 { get; set; }

        [HyperlinkFormat(true)]
        public String Field5 { get; set; }
    }
}
