using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using AutoReview.ExcelParser;
using AutoReview.Structure;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AutoReview.ExcelParser
{
    public class ExcelParser : IExcelParser
    {
        public List<ClassWithScore> GetClassName(string subjectName)
        {
            throw new NotImplementedException();
        }

        public List<StrongSupportClass> GetHighSupportClass()
        {
            throw new NotImplementedException();
        }

        public List<string> GetProfessinalCoreClass()
        {
            throw new NotImplementedException();
        }

        public int GetScore(string subjectName)
        {
            throw new NotImplementedException();
        }

        public int GetTotalScore()
        {
            throw new NotImplementedException();
        }

        public void Init(string path)
        {
            throw new NotImplementedException();
        }
    }
}
