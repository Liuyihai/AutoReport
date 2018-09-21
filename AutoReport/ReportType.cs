using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReport
{
    public static class ReportType
    {
        
    }

    public enum Report_Type
    {
        /// <summary>
        /// 风险分析与定级报告
        /// </summary>
        Risk_SILlevel,
        /// <summary>
        /// SIL分析报告
        /// </summary>
        SIL_analysis,
        /// <summary>
        /// SIL验证报告
        /// </summary>
        SIL_validate,
        /// <summary>
        /// 误停车分析报告
        /// </summary>
        MTTR_analysis
    }
}
