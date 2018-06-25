using AutoReview.Structure;
using NPOI.HWPF;
using NPOI.XWPF.Extractor;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReview.TrainingPlan
{
    class TrainingPlan : ITrainingPlan
    {
        public string textBody { get; private set; }
        public List<ClassWithScore> classWithScoreList { get; private set; }

        public bool FindClassAndScore(string className, int score)
        {
            return classWithScoreList.Where(x => x.ClassName == className && x.Score == score).Count() > 0;
        }

        public void Init(string path)
        {
            classWithScoreList = new List<ClassWithScore>();
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
                    foreach (var t in xwpf.Tables)
                    {
                        int classNameColumnNumber = -1;
                        int classScoreColumnNumber = -1;
                        var cells = t.Rows[0].GetTableCells();
                        int step = 0;
                        foreach(var c in cells)
                        {
                            if (c.GetText().Trim().Contains("课程名称"))
                            {
                                classNameColumnNumber = step;
                            }

                            if (c.GetText().Trim().Contains("学分"))
                            {
                                classScoreColumnNumber = step;
                            }

                            step++;
                        }

                        if(classNameColumnNumber != -1 && classScoreColumnNumber != -1)
                        {
                            step = 0;
                            foreach(var r in t.Rows)
                            {
                                if(step++ == 0)
                                {
                                    continue;
                                }

                                this.classWithScoreList.Add(new ClassWithScore() {
                                    ClassName = r.GetCell(classNameColumnNumber).GetText(),
                                    Score = float.Parse(r.GetCell(classScoreColumnNumber).GetText())
                                });
                            }
                        }
                    }
                }
            }
        }
    }
}
