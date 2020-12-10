﻿using Microsoft.Win32;
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
        public MainWindow()
        {
            InitializeComponent();
        }

        private void chooseTableButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true) {
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

        private void sheetsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // oWorksheet = oWorkbook.Worksheets["ПАРФОМЧУК"];
            Console.WriteLine(sheetsComboBox.SelectedValue);
            oWorksheet = oWorkbook.Worksheets[sheetsComboBox.SelectedValue];
        }

        private void optimizeButton_Click(object sender, RoutedEventArgs e)
        {
            // Reading
            int colNo = oWorksheet.UsedRange.Columns.Count;
            int rowNo = oWorksheet.UsedRange.Rows.Count;
            // Console.WriteLine("Rows: " + (rowNo - 1));
            object[,] array = oWorksheet.UsedRange.Value;
            for (int j = 1; j <= colNo; j++)
            {
                for (int i = 1; i <= rowNo; i++)
                {
                    if (array[i, j] != null)
                        if (array[i, j].ToString() == "Коефіцієнт ЗП місячний ***")
                        {
                            ratiosColumnNo = j;
                            firstRatiosRowNo = i + 1;
                            for (int m = firstRatiosRowNo; m < rowNo; m++)
                            {
                                // Console.WriteLine(array[m, j]);
                                // Console.WriteLine(array[m, j]?.GetType().Name);
                                if (array[m, j]?.GetType().Name == "Double"
                                    && array[m, j + 1] != null)
                                {
                                    // Console.WriteLine(array[m, j]);
                                    ratios.Add(new MonthModel() { RowNo = m, Ratio = (double)array[m, j] });
                                }
                                // ratios.Add((double)array[m, j]);
                            }

                            // set the value back into the range.
                            // oWorksheet.UsedRange.Value = array;
                            goto OUTPUT;
                        }
                }
            }

        // Output
        OUTPUT:
            // Console.WriteLine("Ratios: " + ratios.Count);
            // ratios.ForEach(Console.WriteLine);

            // Optimization
            int count = ratios.Count;
            int tenPercentCount = (int)Math.Round((double)count * 0.1);
            Console.WriteLine("Ten Percent Count: " + tenPercentCount);
            int exclusionsCountLimit = tenPercentCount;
            if (exclusionsCountLimit > 60)
            {
                exclusionsCountLimit = 60;
            }
            Console.WriteLine("Max Exclusions Count Limit: " + exclusionsCountLimit);

            while (exclusionsCountLimit > 0)
            {
                int currentFirstRowNo = ratios.First().RowNo;
                int currentLastRowNo = currentFirstRowNo + exclusionsCountLimit;
                while (currentLastRowNo <= ratios.Last().RowNo)
                {
                    double avgRatioAfterOptimization =
                               ratios.Where(
                                    r => !Enumerable
                                            .Range(currentFirstRowNo, currentLastRowNo)
                                            .Contains(r.RowNo)
                                ).Average(r => r.Ratio);
                    if (avgRatioAfterOptimization > optimalExclusionRange.AvgRatioAfterOptimization)
                    {
                        optimalExclusionRange.FirstRowNo = currentFirstRowNo;
                        optimalExclusionRange.LastRowNo = currentLastRowNo;
                        optimalExclusionRange.AvgRatioAfterOptimization = avgRatioAfterOptimization;
                    }
                    currentFirstRowNo++;
                    currentLastRowNo++;
                }
                exclusionsCountLimit--;
            }

            double avgRatioBeforeOptimization = ratios.Average(r => r.Ratio);
            Console.WriteLine("Average Ratio Before Optimization: " + avgRatioBeforeOptimization);
            Console.WriteLine("Optimal Exclusion Range: " + optimalExclusionRange);

            // Marking
            foreach (int exRowNo in Enumerable.Range(
                    optimalExclusionRange.FirstRowNo,
                    optimalExclusionRange.LastRowNo
                ))
            {
                array[exRowNo, ratiosColumnNo + 2] = "-";
            }
            oWorksheet.UsedRange.Value = array;

            oWorkbook.Save();

            disposeResources();
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
    }
}
