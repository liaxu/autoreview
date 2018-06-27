using AutoReview.Structure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReview.TrainingPlan
{
    interface ITrainingPlan
    {
        void Init(string path);
        Response FindClassAndScore(string className, float score);
    }
}
