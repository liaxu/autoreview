using AutoReview.Structure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AutoReview.ExcelParser
{
    interface IExcelParser
    {
        void Init(string path);
        int GetTotalScore();
        int GetScore(string subjectName);

        List<ClassWithScore> GetClassName(string subjectName);

        List<string> GetProfessinalCoreClass();

        List<StrongSupportClass> GetHighSupportClass();
    }
}
