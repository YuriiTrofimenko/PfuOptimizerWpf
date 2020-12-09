using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PfuOptimizerWpf.Models
{
    class MonthModel
    {
        public int RowNo { get; set; }
        public double Ratio { get; set; }
        public override string ToString()
        {
            return $"MonthModel: [RowNo={RowNo}, Ratio={Ratio}]";
        }
    }
}
