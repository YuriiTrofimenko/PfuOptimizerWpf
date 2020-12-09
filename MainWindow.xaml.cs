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

namespace PfuOptimizerWpf
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<double> ratios = new List<double>();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void chooseTableButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true) {
                Console.WriteLine(openFileDialog.FileName);
                Excel._Application oApp = new Excel.Application();
                // oApp.Visible = true;

                Excel.Workbook oWorkbook = oApp.Workbooks.Open(openFileDialog.FileName);
                Excel.Worksheet oWorksheet = oWorkbook.Worksheets["ПАРФОМЧУК"];

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
                                for (int m = i + 1; m < rowNo; m++)
                                {
                                    /* if (Convert.ToInt32(array[m, j].ToString()) > 50)
                                    {
                                        array[m, j + 1] = "Yes";
                                    } */
                                    // Console.WriteLine(array[m, j]);
                                    // Console.WriteLine(array[m, j]?.GetType().Name);
                                    if (array[m, j]?.GetType().Name == "Double")
                                    {
                                        // Console.WriteLine(array[m, j]);
                                        ratios.Add((double)array[m, j]);
                                    }
                                    /* else {
                                        ratios.Add(0d);
                                    } */
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
                int tenPercentCount = (int) Math.Round((double)count * 0.1);
                Console.WriteLine("Ten Percent Count: " + tenPercentCount);
                int exclusionsCountLimit = tenPercentCount;
                if (exclusionsCountLimit > 60)
                {
                    exclusionsCountLimit = 60;
                }
                Console.WriteLine("Exclusions Count Limit: " + exclusionsCountLimit);

                //oWorkbook.Save();
                oWorkbook.Close();
                oApp.Quit();

                oWorksheet = null;
                oWorkbook = null;
                oApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            // txtEditor.Text = File.ReadAllText(openFileDialog.FileName);
        }
    }
}
