using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReport
{
    /// <summary>
    /// 抽象文档操作父类
    /// </summary>
    public abstract class Operation
    {
        public Operation()
        {

        }
        /// <summary>
        /// 合并文档抽象方法
        /// </summary>
        public abstract void DocMerge(String filename,Form1 form);
    }
}
