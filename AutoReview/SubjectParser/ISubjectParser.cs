using AutoReview.Structure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReview.SubjectParser
{
    interface ISubjectParser
    {
        void Init(string path);
        bool FindClass(string className);
        bool FindClassWithSupportPoint(StrongSupportClass strongSupportClass);
    }
}
