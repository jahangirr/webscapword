using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace webscapword.Models
{

    public class retValue
    {
        public string url { get; set; }
        public List<showValue> showValues { get; set; }
    }

    public class showValue
    {
        public string missingText { get; set; }
    }
}