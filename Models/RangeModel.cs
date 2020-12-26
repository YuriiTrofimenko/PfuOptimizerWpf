using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PfuOptimizerWpf.Models
{
    class RangeModel
    {
        public int FirstRowNo { get; set; }
        public int LastRowNo { get; set; }
        public double AvgRatioAfterProcessing { get; set; }
        public override string ToString()
        {
            return $"ExclusionRangeModel: [FirstRowNo={FirstRowNo}, LastRowNo={LastRowNo}, AvgRatioAfterProcessing={AvgRatioAfterProcessing}]";
        }
    }
}
