using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReview.AnalysisResult
{
    interface IAnalysisResult
    {
        void Init(string name);
        void AddLog(string content);
    }
}
