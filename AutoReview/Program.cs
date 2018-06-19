using AutoReview.ClsNameTest;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReview
{
    class Program
    {
        static void Main(string[] args)
        {
            IClsNameTest ct = new ClsNameRuleTest();
            ct.TestName("高等数学", "数学");
        }
    }
}
