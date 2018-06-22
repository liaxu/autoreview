using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReview.ClsNameTest
{
    class ClsNameRule
    {
        public string SubjectName { get; set; }
        public List<string> AllowRule { get; set; }
        public List<string> DeniedRule { get; set; }

    }
}
