using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AutoReview.ClsNameTest
{
    class ClsNameRuleTest : IClsNameTest
    {
        public bool TestName(string className, string subjectName)
        {
            var text = File.ReadAllText("NameRule.json");
            List<ClsNameRule> nameRuleList = JsonConvert.DeserializeObject<List<ClsNameRule>>(text);
            ClsNameRule rule = nameRuleList.Where(x => x.SubjectName == subjectName).FirstOrDefault();
            if(rule != null)
            {
                foreach(var r in rule.AllowRule)
                {
                    if(r.Trim() != string.Empty)
                    {
                        Regex regex = new Regex(r);
                        if (regex.IsMatch(className))
                        {
                            return true;
                        }
                    }
                }

                foreach (var r in rule.DeniedRule)
                {
                    if (r.Trim() != string.Empty)
                    {
                        Regex regex = new Regex(r);
                        if (regex.IsMatch(className))
                        {
                            return false;
                        }
                    }
                }
            }

            return false;
        }
    }
}
