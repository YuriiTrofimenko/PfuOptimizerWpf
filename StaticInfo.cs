using PfuOptimizerWpf.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PfuOptimizerWpf
{
    class StaticInfo
    {
        public static List<MonthNameModel> GetUkrainianMonths() {
            return new List<MonthNameModel>() {
                new MonthNameModel { No = 1, Month = "Січень" },
                new MonthNameModel { No = 2, Month = "Лютий" },
                new MonthNameModel { No = 3, Month = "Березень" },
                new MonthNameModel { No = 4, Month = "Квітень" },
                new MonthNameModel { No = 5, Month = "Травень" },
                new MonthNameModel { No = 6, Month = "Червень" },
                new MonthNameModel { No = 7, Month = "Липень" },
                new MonthNameModel { No = 8, Month = "Серпень" },
                new MonthNameModel { No = 9, Month = "Вересень" },
                new MonthNameModel { No = 10, Month = "Жовтень" },
                new MonthNameModel { No = 11, Month = "Листопад" },
                new MonthNameModel { No = 12, Month = "Грудень" }
            };
        }
    }
}
