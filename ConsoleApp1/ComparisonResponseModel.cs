using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LiveOptics.DataDomainAutomation;

namespace LiveOptics.DataDomainAutomation
{
    public class ComparisonResponseModel
    {
        public bool Passed { get; set; }

        public string ResponseText { get; set; }
    }
}