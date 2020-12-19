using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PfuOptimizerWpf.Models;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace PfuOptimizerWpf
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private Excel._Application oApp;
        private Excel.Workbook oWorkbook;
        private Excel.Worksheet oWorksheet;
        private bool isDisposed = true;

        private List<MonthModel> ratios = new List<MonthModel>();
        private List<ExclusionRangeModel> exclusionRanges = new List<ExclusionRangeModel>();
        ExclusionRangeModel optimalExclusionRange =
            new ExclusionRangeModel() { AvgRatioAfterOptimization = 0 };
        private int ratiosColumnNo = 0;
        private int firstRatiosRowNo = 0;

        private static readonly string DEFAULT_RATIO_COLUMN_NAME = "Коефіцієнт ЗП місячний ***";
        private string ratioColumnName;
        private string experienceMonthString = "";

        private string windowTitle;
        public MainWindow()
        {
            InitializeComponent();
            ratioColumnNameTextBox.Text = DEFAULT_RATIO_COLUMN_NAME;
        }

        private void chooseTableButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == true)
                {
                    windowTitle = this.Title;
                    this.Title = "Загрузка файла ...";
                    Console.WriteLine(openFileDialog.FileName);
                    oApp = new Excel.Application();
                    this.isDisposed = false;
                    // oApp.Visible = true;

                    oWorkbook = oApp.Workbooks.Open(openFileDialog.FileName);
                    List<string> sheetNames = new List<string>();
                    foreach (Worksheet worksheet in oWorkbook.Worksheets)
                    {
                        sheetNames.Add(worksheet.Name);
                    }
                    sheetsComboBox.ItemsSource = sheetNames;
                }
                // txtEditor.Text = File.ReadAllText(openFileDialog.FileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                disposeResources();
            }
            finally
            {
                this.Title = windowTitle;
            }
        }

        private void sheetsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // oWorksheet = oWorkbook.Worksheets["ПАРФОМЧУК"];
            Console.WriteLine(sheetsComboBox.SelectedValue);
            oWorksheet = oWorkbook.Worksheets[sheetsComboBox.SelectedValue];
        }

        private void optimizeButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                windowTitle = this.Title;
                this.Title = "Попытка оптимизации ...";
                // Reading
                int colNo = oWorksheet.UsedRange.Columns.Count;
                int rowNo = oWorksheet.UsedRange.Rows.Count;
                object[,] array = oWorksheet.UsedRange.Value;
                for (int j = 1; j <= colNo; j++)
                {
                    for (int i = 1; i <= rowNo; i++)
                    {
                        if (array[i, j] != null)
                            if (array[i, j].ToString() == this.ratioColumnName)
                            {
                                ratiosColumnNo = j;
                                firstRatiosRowNo = i + 1;
                                for (int m = firstRatiosRowNo; m < rowNo; m++)
                                {
                                    if (array[m, j]?.GetType().Name == "Double"
                                        && array[m, j + 1] != null)
                                    {
                                        ratios.Add(new MonthModel() { RowNo = m, Ratio = (double)array[m, j] });
                                    }
                                }
                                // set the value back into the range.
                                // oWorksheet.UsedRange.Value = array;
                                goto OUTPUT;
                            }
                    }
                }

            // Output
            OUTPUT:
                // Optimization
                // полное число месяцев работы, в соответствии с числом коэффициентов,
                // считанных из колонки
                int experienceMonthFull = ratios.Count;
                // если поле числа месяцев стажа пусто -
                // принимаем для расчета полное число месяцев работы
                int expirienceMonthComputed =
                    this.experienceMonthString == ""
                        ? experienceMonthFull
                        : Int32.Parse(this.experienceMonthString);
                // берем 10% от всех месяцев стажа
                int tenPercentCount =
                    (int)Math.Round((double)expirienceMonthComputed * 0.1);
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
                    int currentLastRowNo = currentFirstRowNo + exclusionsCountLimit;
                    // пока текущая нижняя граница не совпадет с номером последней строки
                    // в колонке коэффициентов
                    while (currentLastRowNo <= ratios.Last().RowNo)
                    {
                        // вычисляем среднее арифметическое из колонки коэффициентов,
                        // исключая коэффициенты текущего диапазона исключения
                        double avgRatioAfterOptimization =
                                   ratios.Where(
                                        r => !Enumerable
                                                .Range(currentFirstRowNo, currentLastRowNo)
                                                .Contains(r.RowNo)
                                    ).Average(r => r.Ratio);
                        // если получившееся среднее больше среднего из модели оптимального диапазона исключения -
                        // записываем в модель это новое среднее
                        if (avgRatioAfterOptimization > optimalExclusionRange.AvgRatioAfterOptimization)
                        {
                            optimalExclusionRange.FirstRowNo = currentFirstRowNo;
                            optimalExclusionRange.LastRowNo = currentLastRowNo;
                            optimalExclusionRange.AvgRatioAfterOptimization = avgRatioAfterOptimization;
                        }
                        // смещаем диапазон исключения вниз на одну строку
                        currentFirstRowNo++;
                        currentLastRowNo++;

                    }
                    // уменьшаем размер диапазона исключения на одну строку
                    exclusionsCountLimit--;
                }
                // вычисляем среднее арифметическое коэффициентов до оптимизации
                double avgRatioBeforeOptimization = ratios.Average(r => r.Ratio);
                Console.WriteLine("Average Ratio Before Optimization: " + avgRatioBeforeOptimization);
                Console.WriteLine("Optimal Exclusion Range: " + optimalExclusionRange);
                // Marking
                for (int exRowNo = optimalExclusionRange.FirstRowNo; exRowNo <= optimalExclusionRange.LastRowNo; exRowNo++)
                {
                    // рядом с каждым исключенным коэффициентом,
                    // отступив на две ячейки вправо,
                    // отмечаем исключение знаком -
                    Console.WriteLine("exRowNo = " + exRowNo);
                    array[exRowNo, ratiosColumnNo + 2] = "-";
                }
                oWorksheet.UsedRange.Value = array;

                oWorkbook.Save();

                disposeResources();
                resetForm();
            }
            catch (Exception ex)
            {
                // throw;
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                this.Title = windowTitle;
                disposeResources();
            }
        }

        private void resetForm() {
            this.ratioColumnName = DEFAULT_RATIO_COLUMN_NAME;
            this.experienceMonthString = "";
            experienceMonthTextBox.Text = "";
        }

        private void disposeResources() {
            if (!this.isDisposed)
            {
                oWorkbook.Close();
                oApp.Quit();

                oWorksheet = null;
                oWorkbook = null;
                oApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                this.isDisposed = true;
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            disposeResources();
            base.OnClosing(e);
        }

        private void ratioColumnNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Console.WriteLine(ratioColumnNameTextBox.Text);
            this.ratioColumnName = ratioColumnNameTextBox.Text;
        }

        private void experienceMonthTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            this.experienceMonthString = experienceMonthTextBox.Text;
        }
    }
}
