﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using AutoReview.Structure;
using DocumentFormat.OpenXml.Packaging;
using NPOI.HWPF;
using NPOI.XWPF.Extractor;
using NPOI.XWPF.UserModel;

namespace AutoReview.SubjectParser
{
    class SubjectParserCheck : ISubjectParser
    {
        public string textBody { get; set; }
        public bool FindClass(string className)
        {
            Regex regex = null;
            // 课程名称可能存在不同，例如 高等数学A(1)会合并为高等数学A，因此如果找不到课程可以尝试去除小标号
            regex = new Regex(string.Format(".*?{0}.*?((教学大纲)|(课程简介))", className));
            if (regex.IsMatch(textBody) == false)
            {
                regex = new Regex(@"\(\d+\)");
                className = regex.Replace(className, string.Empty);
            }

            className = className.Replace("（", "(").Replace("）", ")");
            regex = new Regex(string.Format(".*?{0}.*?((教学大纲)|(课程简介))", className));
            // todo we may lost a lot of time here
            bool isMatch = regex.IsMatch(textBody);
            return isMatch;
        }

        public bool FindClassWithSupportPoint(string className, string supportPoint)
        {
            // 课程名称可能存在不同，例如 高等数学A(1)会合并为高等数学A，因此如果找不到课程可以尝试去除小标号
            if (textBody.IndexOf(className) == -1)
            {
                Regex regex = new Regex(@"\(\d+\)");
                className = regex.Replace(className, string.Empty);
            }

            int firstPozClassName = textBody.IndexOf(string.Format("《{0}》", className));
            int nextPozMark = textBody.IndexOf("《", firstPozClassName + 1);

            // 跳过目录
            while (nextPozMark != -1 && nextPozMark - firstPozClassName < 500)
            {
                firstPozClassName = textBody.IndexOf(string.Format("《{0}》", className), firstPozClassName + 1);
                nextPozMark = textBody.IndexOf("《", firstPozClassName + 1);
            }

            if (firstPozClassName == -1)
            {
                return false;              
            }

            // 进入正文
            int nextClassSectionPoz = textBody.IndexOf("教学大纲", firstPozClassName + 200);
            string classSection = string.Empty;
            if (nextClassSectionPoz != -1)
            {
                classSection = textBody.Substring(firstPozClassName, nextClassSectionPoz - firstPozClassName);
            }
            else
            {
                classSection = textBody.Substring(firstPozClassName, textBody.Length - firstPozClassName);
            }

            if (classSection.IndexOf(supportPoint) == -1)
            {
                return false;             
            }
            return true;
        }

        public void Init(string path)
        {
            if (path.EndsWith(".doc"))
            {
                HWPFDocument hwpf;
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    hwpf = new HWPFDocument(file);
                }

                this.textBody = hwpf.Text.ToString();
            }
            else if (path.EndsWith(".docx"))
            {
                XWPFDocument xwpf;
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    xwpf = new XWPFDocument(file);
                }

                XWPFWordExtractor ex = new XWPFWordExtractor(xwpf);
                this.textBody = ex.Text;
            }

            textBody = textBody.Replace("（", "(").Replace("）", ")");
        }
    }
}
