using AutoReview.ClsNameTest;
using AutoReview.SubjectParser;
using AutoReview.TrainingPlan;
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

            /*
            ISubjectParser sp = new SubjectParserCheck();
            sp.Init("教学大纲.docx");
            sp.FindClassWithSupportPoint(new Structure.StrongSupportClass() {
                ClassName = "机械制图 A",
                SupportPoint = new List<string>() { "5.1"}
            });
            */

            ITrainingPlan trainingPlan = new TrainingPlan.TrainingPlan();
            trainingPlan.Init("培养方案.docx");
            trainingPlan.FindClassAndScore("高等数学A（1）",5);
        }
    }
}
