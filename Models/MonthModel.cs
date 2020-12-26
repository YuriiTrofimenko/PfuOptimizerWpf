using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PfuOptimizerWpf.Models
{
    class MonthModel
    {
        public int RowNo { get; set; } // номер строки из электронной таблицы
        public double Ratio { get; set; } // коэффициент
        public ushort Year { get; set; } // год, к которому относится месяц стажа
        public string Month { get; set; } // месяц стажа (название)
        public override string ToString()
        {
            return $"MonthModel: [RowNo={RowNo}, Ratio={Ratio}, Year={Year}, Month={Month}]";
        }
    }
}
