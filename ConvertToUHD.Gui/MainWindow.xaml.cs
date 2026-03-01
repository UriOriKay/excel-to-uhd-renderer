using ConvertToUHD.Core;
using Microsoft.Win32;
using System.ComponentModel;
using System.IO;
using System.Windows;
using WindowsAPICodePack.Dialogs;

namespace ConvertToUHD.Gui
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select an Excel file",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                CheckFileExists = true
            };

            if (dialog.ShowDialog() == true)
                InputTextBox.Text = dialog.FileName;
        }

        private void SelectOutput_Click(object sender, RoutedEventArgs e)
        {
            using var dialog = new CommonOpenFileDialog
            {
                Title = "Select output folder",
                IsFolderPicker = true
            };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                OutputTextBox.Text = dialog.FileName;

        }

        private void Convert_Click(object sender, RoutedEventArgs e)
        {

            string inputPath = InputTextBox.Text?.Trim() ?? "";
            string outputDir = OutputTextBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
            {
                StatusText.Text = "Please select a valid .xlsx file.";
                return;
            }

            if (string.IsNullOrWhiteSpace(outputDir))
            {
                StatusText.Text = "Please select an output folder.";
                return;
            }

            try
            {
                Directory.CreateDirectory(outputDir);

                var converter = new ExcelToUhdConverter();
                converter.Convert(inputPath, outputDir, msg => StatusText.Text = msg);

                StatusText.Text = "Done.";
            }
            catch (Exception ex)
            {
                StatusText.Text = "Error: " + ex.Message;
            }
        }
    }
}