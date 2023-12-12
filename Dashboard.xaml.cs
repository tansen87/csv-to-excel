using Microsoft.Win32;
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
using System.Globalization;
using System.Diagnostics;
using System.Security.Policy;
using CsvHelper;
using System.IO;
using CsvHelper.Configuration;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs;

namespace EAT
{
    /// <summary>
    /// Dashboard.xaml 的交互逻辑
    /// </summary>

public partial class Dashboard : UserControl
    {
        public Dashboard()
        {
            InitializeComponent();
        }

        private string[]? selectedFiles;

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "csv file(*.csv)|*.csv|所有文件(*.*)|*.*",
                Multiselect = true
            };
            bool? flag = openFileDialog.ShowDialog();
            bool flag2 = true;
            if (flag.GetValueOrDefault() == flag2 & flag != null)
            {
                string[]? selectedFile = openFileDialog.FileNames;
                if (selectedFile != null && selectedFile.Length != 0)
                {
                    this.TextFilePath.Text = selectedFile[0];
                    this.selectedFiles = selectedFile;
                    return;
                } 
                else
                {
                    this.TextFilePath.Text = "no select file";
                }
            }
        }

        private void BtnGithub_Click(object sender, RoutedEventArgs e)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "https://github.com/tansen87",
                UseShellExecute = true
            };
            Process.Start(psi);
        }

        private int CountRows(string filePath)
        {
            int result;
            using (StreamReader reader = new(filePath))
            {
                using CsvReader csv = new(reader, CultureInfo.InvariantCulture, false);
                int rowCount = 0;
                while (csv.Read())
                {
                    rowCount++;
                }
                result = rowCount;
            }
            return result;
        }

        private string GetDelimiterValue(TextBox delimiter)
        {
            string delimiterValue = string.Empty;
            delimiter.Dispatcher.Invoke(delegate ()
            {
                delimiterValue = delimiter.Text;
            });
            return delimiterValue;
        }

        private string GetColumnsValue(TextBox columns)
        {
            string columnsValue = string.Empty;
            columns.Dispatcher.Invoke(delegate ()
            {
                columnsValue = columns.Text;
            });
            return columnsValue;
        }

        private async void Btnc2x_Click(object sender, RoutedEventArgs e)
        {
            if (selectedFiles == null || selectedFiles.Length == 0)
            {
                display.Text += $"{DateTime.Now.ToString()} => Error:\nNo file open. Please select a file.\n";
                return;
            }
            MessageBox.Show($"{DateTime.Now.ToString()} => Info:\nrunning...");
            for (int i = 0; i < selectedFiles.Length; i++)
            {
                await Task.Run(() =>
                {
                    try
                    {
                        string filePath = selectedFiles[i];
                        int rows = CountRows(filePath);
                        string sep = GetDelimiterValue(delimiter);
                        List<string> floatType = new();
                        string numericColumn = GetColumnsValue(columns);
                        floatType.AddRange(numericColumn.Split(sep));
                        if (rows < 1040000)
                        {
                            List<Dictionary<string, object>> csvData = WriteCsvDataToClassCode(filePath, floatType);
                            var config = new OpenXmlConfiguration()
                            {
                                TableStyles = TableStyles.None
                            };
                            string savePath = System.IO.Path.ChangeExtension(filePath, ".xlsx");
                            MiniExcel.SaveAs(savePath, csvData, overwriteFile: true, configuration: config);
                            display.Dispatcher.Invoke(() =>
                            {
                                display.Text += $"{DateTime.Now.ToString()} => Info:\nsave file: {savePath} {Environment.NewLine}\n";
                            });
                            csvData.Clear();
                            csvData = null;
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        else
                        {
                            display.Dispatcher.Invoke(() =>
                            {
                                display.Text += $"{DateTime.Now.ToString()} => Tips: \n{filePath} more than 1040000.\n";
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        display.Dispatcher.Invoke(() =>
                        {
                            display.Text += $"{DateTime.Now.ToString()} => Error:\n{ex.Message}\n";
                        });
                    }
                });
            }
            MessageBox.Show($"{DateTime.Now.ToString()} => Info:\ndone.");
        }

        private List<Dictionary<string, object>> WriteCsvDataToClassCode(string filePath, List<string> floatType)
        {
            List<Dictionary<string, object>> csvData = new();
            IEnumerable<string> source = File.ReadLines(filePath);
            string sep = this.GetDelimiterValue(this.delimiter);
            string[] propertyNames = source.First<string>().Split(sep, StringSplitOptions.None);
            using (StreamReader reader = new(filePath))
            {
                using (CsvReader csv = new(reader, new CsvConfiguration(CultureInfo.InvariantCulture, null)
                {
                    Delimiter = this.GetDelimiterValue(this.delimiter),
                    HasHeaderRecord = false
                }, false))
                {
                    reader.ReadLine();
                    while (csv.Read())
                    {
                        Dictionary<string, object> data = new();
                        for (int i = 0; i < propertyNames.Length; i++)
                        {
                            string propertyName = propertyNames[i].Trim('"');
                            string propertyValue = csv.GetField(i).Trim('"');
                            if (floatType.Contains(propertyName))
                            {
                                double.TryParse(propertyValue, out double numericValue);
                                data.Add(propertyName, numericValue);
                            }
                            else
                            {
                                data.Add(propertyName, propertyValue);
                            }
                        }
                        csvData.Add(data);
                    }
                }
            }
            return csvData;
        }
    }
}
