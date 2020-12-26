using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using PfuOptimizerWpf.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PfuOptimizerWpf.Io
{
    class IoExcel
    {
        public static List<MonthModel> GetMonthRatios(
            Worksheet oWorksheet,
            string ratioColumnName,
            ref int ratiosColumnNo,
            ref int firstRatiosRowNo
        )
        {
            List<MonthModel> ratios = new List<MonthModel>();
            int colNo = oWorksheet.UsedRange.Columns.Count;
            int rowNo = oWorksheet.UsedRange.Rows.Count;
            object[,] array = oWorksheet.UsedRange.Value;
            // перемещение по всем колонкам j
            for (int j = 1; j <= colNo; j++)
            {
                // перемещение по всем строкам i
                for (int i = 1; i <= rowNo; i++)
                {
                    // если ячейка i, j не пуста
                    if (array[i, j] != null)
                        // если это - колонка коэффициентов
                        if (array[i, j].ToString() == ratioColumnName)
                        {
                            // запоминаем номер колонки коэффициентов
                            ratiosColumnNo = j;
                            // запоминаем номер первой строки с коэффициентом в колонке коэффициентов
                            firstRatiosRowNo = i + 1;
                            // перемещение по всем строкам m колонки коэффициентов
                            for (int m = firstRatiosRowNo; m < rowNo; m++)
                            {
                                // если текущаяячейка не пуста,
                                // содержит число,
                                // и колонка, следующая через одну после текщей, существует
                                if (array[m, j]?.GetType().Name == "Double"
                                    && array[m, j + 1] != null)
                                {
                                    // добавление модели сведений за месяц стажа в список
                                    // Console.WriteLine($"row{m} Y={array[m, 1]} M={array[m, 2]}");
                                    ratios.Add(
                                        new MonthModel()
                                        {
                                            RowNo = m,
                                            Ratio = (double)array[m, j],
                                            Year = UInt16.Parse(array[m, 1].ToString()),
                                            Month = array[m, 2].ToString(),
                                            MinSalaryUkraine = Double.Parse(array[m, 3].ToString()),
                                            AvgSalaryUkraine = Double.Parse(array[m, 4].ToString()),
                                            Income = Double.Parse(array[m, 5].ToString()),
                                            Days = Int32.Parse(array[m, j + 1].ToString())
                                        }
                                    );
                                }
                            }
                            return ratios;
                        }
                }
            }
            return null;
        }

        public static void WriteSelectedMonthRatios(
            _Application oApp,
            string customerName,
            List<MonthModel> selectedRatios
        )
        {
            Workbook worKbooK = oApp.Workbooks.Add(Type.Missing);
            Worksheet newWorKsheeT = (Worksheet)worKbooK.ActiveSheet;
            try
            {
                newWorKsheeT.Name = customerName + " (после обработки)";
            }
            catch (Exception)
            {
                newWorKsheeT.Name = "Фамилия (после обработки)";
            }
            newWorKsheeT.Cells[1, 1] = "Рік";
            newWorKsheeT.Cells[1, 2] = "Місяць";
            newWorKsheeT.Cells[1, 3] = "Мінімальна ЗП по НГ в Україні *";
            newWorKsheeT.Cells[1, 4] = "Середня ЗП по НГ в Україні";
            newWorKsheeT.Cells[1, 5] = "Заробіток, грн **";
            newWorKsheeT.Cells[1, 6] = "Коефіцієнт ЗП місячний ***";
            newWorKsheeT.Cells[1, 7] = "Зараховано до СС, днів *";
            int rowCount = 2;
            foreach (MonthModel ratio in selectedRatios)
            {
                newWorKsheeT.Cells[rowCount, 1] = ratio.Year;
                newWorKsheeT.Cells[rowCount, 2] = ratio.Month;
                newWorKsheeT.Cells[rowCount, 3] = ratio.MinSalaryUkraine;
                newWorKsheeT.Cells[rowCount, 4] = ratio.AvgSalaryUkraine;
                newWorKsheeT.Cells[rowCount, 5] = ratio.Income;
                newWorKsheeT.Cells[rowCount, 6] = ratio.Ratio;
                newWorKsheeT.Cells[rowCount, 7] = ratio.Days;
                rowCount++;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = newWorKsheeT.Name + ".xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    worKbooK.SaveAs(saveFileDialog.FileName);
                }
                catch (Exception)
                {
                    MessageBox.Show("Невозможно сохранить файл. Возможно, он уже существует и открыт.");
                }
            }
            worKbooK.Close();
        }
    }
}
