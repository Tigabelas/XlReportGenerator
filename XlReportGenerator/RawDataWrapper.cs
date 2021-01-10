using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlReportGenerator
{
    public class RawDataWrapper
    {
        public String Cell { get; set; }
        public String Value { get; set; }
        public Int16 Type { get; set; }  //1= Numeric, 2=Text, 5=Date
    }
}
