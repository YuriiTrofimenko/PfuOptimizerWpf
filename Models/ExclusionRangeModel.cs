using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PfuOptimizerWpf.Models
{
    class ExclusionRangeModel
    {
        public int FirstRowNo { get; set; }
        public int LastRowNo { get; set; }
        public double AvgRatioAfterOptimization { get; set; }
    }
}
