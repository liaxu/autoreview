﻿using AutoReview.ClsNameTest;
using AutoReview.Structure;
using AutoReview.SubjectParser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReview.Rule
{
    class Rule2 : IRule
    {
        public void Check()
        {
            ZipFileManager.ZipFileManager zfm = new ZipFileManager.ZipFileManager();
            ClsNameRuleTest cnt = new ClsNameRuleTest();
            ExcelParser.ExcelParser ep = new ExcelParser.ExcelParser();
            string zipFilePath = "";
            try
            {
                zipFilePath = zfm.UnZip(@"d:\报告.zip");
            }
            catch(Exception x)
            {
                AnalysisResult.AnalysisLog.AddLog("【系统】"+x.Message);
            }
           
            SubjectParserCheck subjectParserCheck = new SubjectParserCheck();
            // 找到教学大纲并进行初始化
            var trainingPlan = zfm.FindFile(zipFilePath, "*教学大纲.*");
            if(trainingPlan!=null && trainingPlan.Count > 0)
            {
                subjectParserCheck.Init(trainingPlan.FirstOrDefault());
            }
            else
            {
                AnalysisResult.AnalysisLog.AddLog("【系统】无法找到教学大纲");
            }
            // 找到课程体系对毕业要求指标点的支撑关系表excel
            var supportClassFiles=zfm.FindFile(zipFilePath, "*支撑关系表.*");
            List<StrongSupportClass> strongSupportClassList = null;
            if (supportClassFiles!=null && supportClassFiles.Count > 0)
            {
                ep.Init(supportClassFiles.FirstOrDefault());
                strongSupportClassList= ep.GetHighSupportClass();
            }
            else
            {
                AnalysisResult.AnalysisLog.AddLog("【系统】无法找到课程体系对毕业要求的支持关系表");
            }
            // 找到强支撑课程需要在课程大纲中出现，并且对应相应的支撑点。
            Response response = null;
            if(strongSupportClassList != null && strongSupportClassList.Count > 0)
            {
                foreach(var strongSupportClass in strongSupportClassList)
                {
                    response=subjectParserCheck.FindClassWithSupportPoint(strongSupportClass);
                    if (response.ReturnCode==1)
                    {
                        AnalysisResult.AnalysisLog.AddLog(response.Message);
                    }
                }
            }

        }
    }
}
