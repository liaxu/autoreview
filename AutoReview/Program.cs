using AutoReview.ClsNameTest;
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
            IClsNameTest ct = new ClsNameRuleTest();
            ct.TestName("高等数学", "数学");

            ISubjectParser sp = new SubjectParserCheck();
            sp.Init("教学大纲.docx");
            sp.FindClass("高脚本");
        }
    }
}
