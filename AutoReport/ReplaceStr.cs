using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Spire.Doc;

namespace AutoReport
{
    class ReplaceStr
    {
        public Document Replace(Document doc ,string[] text)
        {
            //公司
            List<Regex> reg_copm = new List<Regex>();
            reg_copm.Add(new Regex(@"浙江石油化工有限公司"));
            reg_copm.Add(new Regex(@"浙江石化公司"));
            reg_copm.Add(new Regex(@"浙江石化"));
            reg_copm.Add(new Regex(@"浙石化公司"));
            reg_copm.Add(new Regex(@"浙石化"));
            foreach (Regex reg in reg_copm)
            {
                doc.Replace(reg, text[0]);
            }
            //项目
            doc.Replace(new Regex("4000万吨/年炼化一体化项目一期工程\n装置SIL分析及部分装置HAZOP分析项目"), text[1]);
            //装置
            reg_copm.Clear();
            reg_copm.Add(new Regex("450万吨/年重油催化裂化装置"));
            reg_copm.Add(new Regex("重油催化裂化装置"));
            reg_copm.Add(new Regex("催化裂化装置"));
            foreach (Regex reg in reg_copm)
            {
                doc.Replace(reg, text[2]);
            }

            return doc;
        }
    }
}
