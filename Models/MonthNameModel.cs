using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PfuOptimizerWpf.Models
{
    class MonthNameModel
    {
        public ushort No { get; set; } // номер месяца в году
        public string Month { get; set; } // название месяца
        public override string ToString()
        {
            return $"MonthModel: [No={No}, Month={Month}]";
        }
    }
}
