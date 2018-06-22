using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using AutoReview.Structure;
using DocumentFormat.OpenXml.Packaging;

namespace AutoReview.SubjectParser
{
    class SubjectParserCheck : ISubjectParser
    {
        public string textBody { get; set; }
        public bool FindClass(string className)
        {
            Regex regex = new Regex(string.Format(".*?{0}.*?((教学大纲)|(课程简介))",className));
            // todo we may lost a lot of time here
            bool isMatch = regex.IsMatch(textBody);
            return isMatch;
        }

        public bool FindClassWithSupportPoint(StrongSupportClass strongSupportClass)
        {
            throw new NotImplementedException();
        }

        public void Init(string path)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, false))
            {
                this.textBody = wordDocument.MainDocumentPart.Document.Body.InnerText;
            }
        }
    }
}
