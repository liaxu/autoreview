using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReview.AnalysisResult
{
    static class AnalysisLog
    {
        private static Guid LogId { get; set; }
        public static void AddLog(string content)
        {
            if(LogId == Guid.Empty)
            {
                LogId = Guid.NewGuid();
            }

            if (File.Exists(string.Format("{0}.txt", LogId)) == false)
            {
                File.AppendAllLines(string.Format("{0}.txt", LogId), new string[] { "Log start" });
                Console.WriteLine( "Log start");
            }

            File.AppendAllLines(string.Format("{0}.txt", LogId), new string[] {
                    string.Format("{0}  {1}",DateTime.Now, content)
                });

            Console.WriteLine(string.Format("{0}  {1}", DateTime.Now, content));
        }
    }
}
