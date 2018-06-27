using AutoReview.ClsNameTest;
using AutoReview.Rule;
using AutoReview.SubjectParser;
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
            //Rule1 r1 = new Rule1();
            //r1.Check();
            Rule2 r2 = new Rule2();
            r2.Check();
        }
    }
}
