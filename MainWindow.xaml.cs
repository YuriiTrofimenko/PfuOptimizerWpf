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
using PfuOptimizerWpf.Io;
using PfuOptimizerWpf.Processors;

namespace PfuOptimizerWpf
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private _Application oApp;
        private Workbook oWorkbook;
        private Worksheet oWorksheet;
        private bool isDisposed = true;

        // private List<MonthModel> ratios;
        private List<RangeModel> exclusionRanges = new List<RangeModel>();
        /* RangeModel optimalExclusionRange =
            new RangeModel() { AvgRatioAfterProcessing = 0 }; */
        private int ratiosColumnNo = 0;
        private int firstRatiosRowNo = 0;

        private static readonly string DEFAULT_RATIO_COLUMN_NAME = "Коефіцієнт ЗП місячний ***";
        private string ratioColumnName;
        private string experienceMonthString = "";

        private string windowTitle;
        public MainWindow()
        {
            InitializeComponent();
            resetForm();
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
                    if (oApp == null)
                    {
                        oApp = new Excel.Application();
                        this.isDisposed = false;
                    }
                    // oApp.Visible = true;
                    oWorkbook = oApp.Workbooks.Open(openFileDialog.FileName);
                    List<string> sheetNames = new List<string>();
                    foreach (Worksheet worksheet in oWorkbook.Worksheets)
                    {
                        sheetNames.Add(worksheet.Name);
                    }
                    sheetsComboBox.ItemsSource = sheetNames;
                    sheetsComboBox.IsEnabled = true;
                    ratioColumnNameTextBox.IsEnabled = true;
                    experienceMonthTextBox.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                DisposeAppResources();
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
            if (sheetsComboBox.SelectedValue != null)
            {
                oWorksheet = oWorkbook.Worksheets[sheetsComboBox.SelectedValue];
                optimizeButton.IsEnabled = true;
            }
        }

        private void optimizeButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                windowTitle = this.Title;
                this.Title = "Попытка оптимизации ...";
                // чтение списка коэффициентов по месяцам трудового стажа
                // из электронной таблицы Excel
                List<MonthModel> ratios = IoExcel.GetMonthRatios(
                    oWorksheet,
                    ratioColumnName,
                    ref ratiosColumnNo,
                    ref firstRatiosRowNo
                );
                CloseTable();
                if (ratios != null)
                {
                    // Отделение списка моделей месяцев стажа, у которых год меньше 2000,
                    // или год равен 2000, и месяц - из множества
                    // Січень, Лютий, Березень, Квітень, Травень, Червень
                    List<MonthModel> oldRatios = new List<MonthModel>();
                    List<MonthModel> newRatios = new List<MonthModel>();
                    foreach (MonthModel ratio in ratios)
                    {
                        if (
                            ratio.Year < 2000
                            || (ratio.Year == 2000
                                && new List<string>() {
                                    "Січень", "Лютий",
                                    "Березень", "Квітень", "Травень",
                                    "Червень"
                                }.Contains(ratio.Month)))
                        {
                            oldRatios.Add(ratio);
                        }
                        else
                        {
                            newRatios.Add(ratio);
                        }
                    }

                    /* отбор наиболее выгодного диапазона месяцев стажа
                     из списка месяцев до 1 июля 2000 года */

                    oldRatios = Selector.SelectMaxOldRange(oldRatios);
                    // объединение списка старых и новых месяцев стажа
                    ratios = oldRatios;
                    ratios.AddRange(newRatios);
                    // перенумерация списка, начиная с 1
                    int rowNum = 1;
                    foreach (MonthModel ratio in ratios)
                    {
                        ratio.RowNo = rowNum;
                        rowNum++;
                    }

                    /* оптимизация (исключение диапазона наиболее невыгодных месяцев стажа) */

                    // полное число месяцев работы, в соответствии с числом коэффициентов,
                    // считанных из колонки
                    int experienceMonthFull = ratios.Count;
                    // если поле числа месяцев стажа пусто -
                    // принимаем для расчета полное число месяцев работы
                    int expirienceMonthComputed =
                            experienceMonthString == ""
                                ? experienceMonthFull
                                : Int32.Parse(experienceMonthString);
                    // оптимизация
                    List <MonthModel> selectedRatios = Selector.ExcludeMinRange(ratios, expirienceMonthComputed);
                    // среднее арифметическое коэффициентов до оптимизации
                    double avgRatioBeforeOptimization = ratios.Average(r => r.Ratio);
                    Console.WriteLine("Average Ratio Before Optimization: " + avgRatioBeforeOptimization);

                    // создание файла электронной таблицы
                    // с результатами отбора месяцев стажа
                    // для текущего клиента
                    IoExcel.WriteSelectedMonthRatios(
                        oApp, sheetsComboBox.SelectedValue.ToString(), selectedRatios
                    );

                    // Marking
                    /* for (int exRowNo = optimalExclusionRange.FirstRowNo; exRowNo <= optimalExclusionRange.LastRowNo; exRowNo++)
                    {
                        // рядом с каждым исключенным коэффициентом,
                        // отступив на две ячейки вправо,
                        // отмечаем исключение знаком -
                        Console.WriteLine("exRowNo = " + exRowNo);
                        array[exRowNo, ratiosColumnNo + 2] = "-";
                    }
                    oWorksheet.UsedRange.Value = array; */

                    // oWorkbook.Save();
                    resetForm();
                }
                else {
                    throw new Exception("Невозможно получить данные о коэффициентах");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                DisposeAppResources();
            }
            finally
            {
                this.Title = windowTitle;
            }
        }

        private void resetForm() {
            // this.ratioColumnName = DEFAULT_RATIO_COLUMN_NAME;
            ratioColumnNameTextBox.Text = DEFAULT_RATIO_COLUMN_NAME;
            this.experienceMonthString = "";
            experienceMonthTextBox.Text = "";
            sheetsComboBox.ItemsSource = null;
            sheetsComboBox.IsEnabled = false;
            ratioColumnNameTextBox.IsEnabled = false;
            experienceMonthTextBox.IsEnabled = false;
            optimizeButton.IsEnabled = false;
        }

        private void CloseTable() {
            oWorkbook.Close();
            oWorksheet = null;
            oWorkbook = null;
        }

        private void DisposeAppResources() {
            if (!this.isDisposed)
            {
                oApp.Quit();
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
            DisposeAppResources();
            base.OnClosing(e);
        }

        private void ratioColumnNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            this.ratioColumnName = ratioColumnNameTextBox.Text;
        }

        private void experienceMonthTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            this.experienceMonthString = experienceMonthTextBox.Text;
        }
    }
}
