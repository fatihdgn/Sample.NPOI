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
using MahApps.Metro.Controls;
using Microsoft.Win32;
using NPOI.Extension;
using System.IO;

namespace Sample.NPOI.Client
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dg.ItemsSource = new List<ItemModel>();
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "Excel file|*.xlsx",
                Title = "Choose a path to open an excel file."
            };

            if (ofd.ShowDialog() == true)
            {
                dg.ItemsSource = Excel.Load<ItemModel>(ofd.FileName);

                MessageBox.Show("File opened successfully.");
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                OverwritePrompt = true,
                AddExtension = true,
                DefaultExt = ".xlsx",
                Filter = "Excel file|*.xlsx",
                Title = "Choose a path to save the excel file."
            };

            if(sfd.ShowDialog() == true)
            {
                if (File.Exists(sfd.FileName))
                {
                    File.Delete(sfd.FileName);
                }

                dg.ItemsSource.Cast<ItemModel>().ToExcel(sfd.FileName);

                MessageBox.Show("File saved successfully.");
            }
        }
    }
}
