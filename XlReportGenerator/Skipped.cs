using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlReportGenerator
{
    [AttributeUsage(AttributeTargets.All, Inherited = false, AllowMultiple = true)]
    public sealed class Skipped : Attribute
    {
        public String SkippedFor { get; set; }

        public Skipped(String skippedFor="")
        {
            this.SkippedFor = skippedFor;
        }

        public Boolean IsSkipped(String sheetName) {
            Boolean result = true;

            if (String.IsNullOrWhiteSpace(this.SkippedFor))
            {
                result = false;
            }
            else
            {
                result = this.SkippedFor.Contains(sheetName);
            }
            

            return result;
        }
    }
}
