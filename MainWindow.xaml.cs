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

namespace EAT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Frame frameCsv2xlsx = new() { Content = new Dashboard() }; 
        Frame frameGetSheetName = new() { Content = new Page1GetSheetName() };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Btnc2x_Click(object sender, RoutedEventArgs e)
        {
            contentcon.Content = frameCsv2xlsx;
            Btnc2x.Click += (sender, e) => {
                Btnc2x.Foreground = Brushes.Green;
                BtnSheet.Foreground = Brushes.Black;
            };
        }

        private void BtnSheet_Click(object sender, RoutedEventArgs e)
        {
            contentcon.Content = frameGetSheetName;
            BtnSheet.Click += (sender, e) => {
                BtnSheet.Foreground = Brushes.Green;
                Btnc2x.Foreground = Brushes.Black;
            };
        }

        private void MainWindow_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            RenderPages.Children.Clear();
            RenderPages.Children.Add(new Dashboard());
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void BtnScale_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.WindowState = WindowState.Minimized;
        }
    }
}
