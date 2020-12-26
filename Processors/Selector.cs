using PfuOptimizerWpf.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PfuOptimizerWpf.Processors
{
    class Selector
    {
        /* 
         * За бажанням пенсіонера
         * та за умови підтвердження довідки про заробітну плату первинними документами
         * або в разі, якщо страховий стаж починаючи з 1 липня 2000 року становить менше 60 місяців,
         * для обчислення пенсії також враховується заробітна плата (дохід)
         * за будь-які 60 календарних місяців страхового стажу підряд
         * по 30 червня 2000 року незалежно від перерв.
         */
        public static List<MonthModel> SelectMaxOldRange(List<MonthModel> oldRatios) {
            List<MonthModel> maxOldRatios = new List<MonthModel>();
            const byte selectionsCountLimit = 60;
            if (oldRatios.Count > selectionsCountLimit)
            {
                RangeModel optimalSelectionRange =
                    new RangeModel() { AvgRatioAfterProcessing = 0 };
                // номер строки верхней границы первого диапазона -
                // из номера первой строки, содержащей коэффициент 
                int currentFirstRowNo = oldRatios.First().RowNo;
                // номер строки нижней границы первого диапазона -
                // сумма номера строки верхней границы
                // и максимально допустимого числа исключений
                int currentLastRowNo = currentFirstRowNo + selectionsCountLimit - 1;
                // пока текущая нижняя граница не совпадет с номером последней строки
                // в колонке коэффициентов
                while (currentLastRowNo <= oldRatios.Last().RowNo)
                {
                    // вычисляем среднее арифметическое из колонки коэффициентов,
                    // учитывая каждый раз только коэффициенты текущего диапазона выборки
                    double avgRatioAfterSelection =
                               oldRatios.Where(
                                    r => (r.RowNo >= currentFirstRowNo
                                            && r.RowNo <= currentLastRowNo)
                                ).Average(r => r.Ratio);
                    // если получившееся среднее больше среднего из модели оптимального диапазона выборки -
                    // записываем в модель это новое среднее
                    if (avgRatioAfterSelection > optimalSelectionRange.AvgRatioAfterProcessing)
                    {
                        optimalSelectionRange.FirstRowNo = currentFirstRowNo;
                        optimalSelectionRange.LastRowNo = currentLastRowNo;
                        optimalSelectionRange.AvgRatioAfterProcessing = avgRatioAfterSelection;
                    }
                    // смещаем диапазон выборки вниз на одну строку
                    currentFirstRowNo++;
                    currentLastRowNo++;

                }
                Console.WriteLine("Optimal Selection Range: " + optimalSelectionRange);
                // отбираем в результирующий список только модели месяцев диапазона
                // с наибольшим средним коэффииентом
                maxOldRatios = oldRatios.Where(
                                    r => (r.RowNo >= optimalSelectionRange.FirstRowNo
                                            && r.RowNo <= optimalSelectionRange.LastRowNo)
                                ).ToList();
            }
            return maxOldRatios;
        }

        /* 
         * З періоду, за який враховується зарплата для обчислення пенсії,
         * Закон дає змогу виключити періоди до 60 календарних місяців страхового стажу
         * з урахуванням будь-яких періодів одержання допомоги по безробіттю незалежно від перерв,
         * і будь-якого періоду страхового стажу підряд за умови,
         * що зазначені періоди в сумі становлять не більше як 10 % тривалості страхового стажу,
         * врахованого в одинарному розмірі.
         * Із підрахунку виключають тільки невигідні коефіцієнти заробітку.
         */
        public static List<MonthModel> ExcludeMinRange(List<MonthModel> ratios, int expirienceMonthComputed)
        {
            RangeModel optimalExclusionRange =
                new RangeModel() { AvgRatioAfterProcessing = 0 };
            // берем 10% от всех месяцев стажа
            int tenPercentCount =
                (int)Math.Round(expirienceMonthComputed * 0.1);
            Console.WriteLine("Ten Percent Count: " + tenPercentCount);
            // если число месяцев для исключения больше 60 -
            // устанавливаем вместо него максимально допустимое число 60
            int exclusionsCountLimit = tenPercentCount;
            if (exclusionsCountLimit > 60)
            {
                exclusionsCountLimit = 60;
            }
            Console.WriteLine("Max Exclusions Count Limit: " + exclusionsCountLimit);
            // пока текущее число месяцев для исключения больше 0 -
            // выполняем проходы сверху вниз по всему множеству коэффициентов,
            // вычисляя для всех возможных диапазонов исключения
            // среднее арифметическое из колонки коэффициентов
            while (exclusionsCountLimit > 0)
            {
                // номер строки верхней границы первого диапазона текущей размерности -
                // из номера первой строки, содержащей коэффициент 
                int currentFirstRowNo = ratios.First().RowNo;
                // номер строки нижней границы первого диапазона текущей размерности -
                // сумма номера строки верхней границы
                // и максимально допустимого числа исключений
                int currentLastRowNo = currentFirstRowNo + exclusionsCountLimit - 1;
                // пока текущая нижняя граница не совпадет с номером последней строки
                // в колонке коэффициентов
                while (currentLastRowNo <= ratios.Last().RowNo)
                {
                    // вычисляем среднее арифметическое из колонки коэффициентов,
                    // исключая коэффициенты текущего диапазона исключения
                    double avgRatioAfterOptimization =
                                ratios.Where(
                                    r => (r.RowNo < currentFirstRowNo
                                            || r.RowNo > currentLastRowNo)
                                ).Average(r => r.Ratio);
                    // если получившееся среднее больше среднего из модели оптимального диапазона исключения -
                    // записываем в модель это новое среднее
                    if (avgRatioAfterOptimization > optimalExclusionRange.AvgRatioAfterProcessing)
                    {
                        optimalExclusionRange.FirstRowNo = currentFirstRowNo;
                        optimalExclusionRange.LastRowNo = currentLastRowNo;
                        optimalExclusionRange.AvgRatioAfterProcessing = avgRatioAfterOptimization;
                    }
                    // смещаем диапазон исключения вниз на одну строку
                    currentFirstRowNo++;
                    currentLastRowNo++;

                }
                // уменьшаем размер диапазона исключения на одну строку
                exclusionsCountLimit--;
            }
            Console.WriteLine("Optimal Exclusion Range: " + optimalExclusionRange);
            return ratios.Where(r => (r.RowNo < optimalExclusionRange.FirstRowNo
                                            || r.RowNo > optimalExclusionRange.LastRowNo)
                                ).ToList();
        }
    }
}
