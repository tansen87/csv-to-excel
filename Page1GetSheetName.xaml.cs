using Microsoft.Win32;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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

namespace EAT
{
    /// <summary>
    /// Page1GetSheetName.xaml 的交互逻辑
    /// </summary>
    public partial class Page1GetSheetName : UserControl
    {
        public Page1GetSheetName()
        {
            InitializeComponent();
        }

        private string[]? selectedFiles;
        private void BtnOpenExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new()
            {
                Filter = "excel file(*.xlsx)|*.xlsx|所有文件(*.*)|*.*",
                Multiselect = true
            };
            bool? flag = openFileDialog.ShowDialog();
            bool flag2 = true;
            if (flag.GetValueOrDefault() == flag2 & flag != null)
            {
                string[]? selectedFile = openFileDialog.FileNames;
                if (selectedFile != null && selectedFile.Length != 0)
                {
                    this.ExcelFilePath.Text = selectedFile[0];
                    this.selectedFiles = selectedFile;
                    return;
                }
                else
                {
                    this.ExcelFilePath.Text = "no select file";
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

        private async void BtnSheet_Click(object sender, RoutedEventArgs e)
        {
            if (selectedFiles == null || selectedFiles.Length == 0)
            {
                display.Text += $"{DateTime.Now.ToString()} => Error:\nNo file open. Please select a file.\n";
                return;
            }
            await Task.Run(() =>
            {
                try
                {
                    var sheetnames = MiniExcel.GetSheetNames(selectedFiles[0]);
                    var sheetname = string.Join("/", sheetnames);
                    string filename = System.IO.Path.GetFileName(selectedFiles[0]);
                    int dashCount = 68;
                    string dashes = new('-', dashCount);
                    display.Dispatcher.Invoke(() =>
                    {
                        display.Text += $"{filename} - SheetName\n{sheetname}\n{dashes}\n";
                    });
                }
                catch (Exception ex)
                {
                    display.Dispatcher.Invoke(() =>
                    {
                        display.Text += $"{DateTime.Now.ToString()} => Error:\n{ex}\n";
                    });
                }
            });
        }
    }
}
