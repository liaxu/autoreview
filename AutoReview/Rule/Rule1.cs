using AutoReview.AnalysisResult;
using AutoReview.ClsNameTest;
using AutoReview.Structure;
using AutoReview.SubjectParser;
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
            List<string> filePaths = zfm.FindFile(zipFilePath, "*知识领域覆盖表*.*");
            foreach (var path in filePaths)
            {
                ep.Init(path);
                // 数学+自然科学类占总分不得低于15%
                int score = ep.GetScore("数学类") + ep.GetScore("自然科学类");
                int totalScore = ep.GetTotalScore();
                if ((double)score / (double)totalScore < 0.15)
                {
                    AnalysisLog.AddLog(string.Format("【期望】数学+自然科学类占总分不得低于15%   【实际】{0:P2}", (double)score / (double)totalScore));
                }

                // 数学类课程名包含或者不得包含规则
                List<ClassWithScore> classList = ep.GetClassName("数学类");
                foreach (var c in classList)
                {
                    bool ret = cnt.TestName(c.ClassName, "数学类");
                    if (ret == false)
                    {
                        AnalysisLog.AddLog(string.Format("【期望】数学类课程名包含或者不得包含规则   【实际】数学类课程名包含{0}", c.ClassName));
                    }
                }

                // 需存在教学大纲
                List<string> subjectFilePaths = zfm.FindFile(zipFilePath, "*教学大纲.*");
                SubjectParserCheck sp = new SubjectParserCheck();
                if (subjectFilePaths.Count() == 0)
                {
                    AnalysisLog.AddLog("【系统】无法找到教学大纲");
                }
                else
                {
                    string subjectFilePath = subjectFilePaths[0];
                    sp.Init(subjectFilePath);
                }

                // 专业核心课在课程大纲出出现
                List<string> professionalCoreClassList = ep.GetProfessinalCoreClass();
                if (subjectFilePaths.Count() > 0)
                {
                    
                    foreach (var i in professionalCoreClassList)
                    {
                        var ret = sp.FindClass(i);
                        if (ret == false)
                        {
                            AnalysisLog.AddLog(string.Format("【期望】专业核心课在教学大纲出现  【实际】课程{0}没有出现", i));
                        }
                    }
                }

                // 需存在培养方案
                List<string> trainingFilePaths = zfm.FindFile(zipFilePath, "*培养方案.*");
                TrainingPlan.TrainingPlan tp = new TrainingPlan.TrainingPlan();
                if (trainingFilePaths.Count() == 0)
                {
                    AnalysisLog.AddLog("【系统】无法找到培养方案");
                }
                else
                {
                    string trainingFilePath = trainingFilePaths[0];
                    tp.Init(trainingFilePath);
                }

                // 课程和学分在培养方案中可查询到
                List<ClassWithScore> allClassWithScore = new List<ClassWithScore>();
                allClassWithScore.Concat(ep.GetClassName("数学类"));
                foreach (var i in allClassWithScore)
                {
                    var response = tp.FindClassAndScore(i.ClassName, i.Score);
                    if (response.ReturnCode == 1)
                    {
                        AnalysisLog.AddLog(response.Message);
                    }
                }
            }
        }
    }
}
