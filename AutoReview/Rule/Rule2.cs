using AutoReview.ClsNameTest;
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
            string zipFilePath = zfm.UnZip(@"d:\报告.zip");

            // 找到课程体系对毕业要求指标点的支撑关系表excel

            // 找到强支撑课程需要在课程大纲中出现，并且对应相应的支撑点。
        }
    }
}
