using AutoReview.AnalysisResult;
using AutoReview.ClsNameTest;
using AutoReview.Structure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReview.Rule
{
    class Rule1 : IRule
    {
        public void Check()
        {
            ZipFileManager.ZipFileManager zfm = new ZipFileManager.ZipFileManager();
            ClsNameRuleTest cnt = new ClsNameRuleTest();
            ExcelParser.ExcelParser ep = new ExcelParser.ExcelParser();
            string zipFilePath = zfm.UnZip(@"d:\报告.zip");
            List<string> filePaths = zfm.FindFile(zipFilePath, "知识领域覆盖表");
            foreach(var path in filePaths)
            {
                ep.Init(path);
                // 数学+自然科学类占总分不得低于15%
                int score = ep.GetScore("数学类") + ep.GetScore("自然科学类");
                int totalScore = ep.GetTotalScore();
                if((double)score/(double)totalScore < 0.15)
                {
                    AnalysisLog.AddLog(string.Format("【期望】数学+自然科学类占总分不得低于15%   【实际】{0}", (double)score / (double)totalScore));
                }

                // 数学类课程名包含或者不得包含规则
                List<ClassWithScore> classList = ep.GetClassName("数学类");
                foreach(var c in classList)
                {
                    bool ret = cnt.TestName(c.ClassName, "数学类");
                    if(ret == false)
                    {
                        AnalysisLog.AddLog(string.Format("【期望】数学类课程名包含或者不得包含规则   【实际】数学类课程名包含{0}", c.ClassName));
                    }
                }

                // 专业核心课在课程大纲出出现
                List<string> professionalCoreClassList = ep.GetProfessinalCoreClass();

                // 课程和学分在培养方案中可查询到

            }
        }
    }
}
