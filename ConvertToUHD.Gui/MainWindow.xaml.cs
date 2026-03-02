using ConvertToUHD.Core;
using Microsoft.Win32;
using System.IO;
using System.Windows;
using WindowsAPICodePack.Dialogs;

namespace ConvertToUHD.Gui
{
    /// <summary>
    /// Main application window for the ConvertToUHD GUI.
    /// 
    /// Provides:
    /// - Drag & drop support for Excel files
    /// - Multi-file processing queue
    /// - Output folder selection
    /// - Progress reporting during conversion
    /// </summary>
    /// /// <remarks>
    /// This window serves as the presentation layer.
    /// Business logic is delegated to the Core library.
    /// </remarks>
    public partial class MainWindow : Window
    {
        private readonly List<string> _files = new();

        /// <summary>
        /// Initializes the main application window.
        /// 
        /// Loads UI components defined in XAML and
        /// initializes the file queue display.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            RefreshFileList();
        }
        /// <summary>
        /// Initializes the main application window.
        /// </summary>
        /// <remarks>
        /// - Loads all UI components defined in XAML.
        /// - Initializes the file queue display.
        /// - Ensures the UI reflects the current internal state.
        /// </remarks>
        private void RefreshFileList()
        {
            FileListBox.ItemsSource = null;
            FileListBox.ItemsSource = _files;
            StatusText.Text = $"{_files.Count} file(s) in queue.";
        }

        /// <summary>
        /// Filters a collection of file paths and returns only existing
        /// Excel (.xlsx) files.
        /// </summary>
        /// <param name="paths">
        /// Collection of file system paths to validate.
        /// </param>
        /// <returns>
        /// An enumerable containing only valid, existing .xlsx files.
        /// </returns>
        /// <remarks>
        /// - File existence is verified before extension check.
        /// - Extension comparison is case-insensitive.
        /// - Uses deferred execution (LINQ).
        /// </remarks>
        private static IEnumerable<string> FilterExcelFiles(IEnumerable<string> paths)
        {
            return paths
                .Where(p => File.Exists(p))
                .Where(p => string.Equals(Path.GetExtension(p), ".xlsx", StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Adds valid Excel files to the internal processing queue.
        /// </summary>
        /// <param name="paths">
        /// Collection of file paths to evaluate and add.
        /// </param>
        /// <remarks>
        /// - Only existing .xlsx files are accepted.
        /// - Duplicate entries are prevented (case-insensitive).
        /// - UI is refreshed after updating the queue.
        /// </remarks>
        private void AddToQueue(IEnumerable<string> paths)
        {
            foreach ( var p in FilterExcelFiles(paths))
            {
                if (!_files.Contains(p, StringComparer.OrdinalIgnoreCase))
                    _files.Add(p);
            }

            RefreshFileList();
        }

        /// <summary>
        /// Opens a folder selection dialog and assigns the selected
        /// directory as output path.
        /// </summary>
        /// <param name="sender">Event source.</param>
        /// <param name="e">Routed event data.</param>
        /// <remarks>
        /// Uses Windows API Code Pack to provide a native
        /// folder picker dialog.
        /// </remarks>
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

        /// <summary>
        /// Opens a file selection dialog and adds selected Excel files
        /// to the processing queue.
        /// </summary>
        /// <param name="sender">Event source.</param>
        /// <param name="e">Routed event data.</param>
        /// <remarks>
        /// - Allows multiple file selection.
        /// - File existence is validated automatically by the dialog.
        /// - Further validation (extension check, duplicates) is handled
        ///   by <see cref="AddToQueue"/>.
        /// </remarks>
        private void AddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select an Excel file",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Multiselect = true,
                CheckFileExists = true
            };

            if (dialog.ShowDialog() == true)
                AddToQueue(dialog.FileNames);
        }

        /// <summary>
        /// Removes the currently selected files from the processing queue.
        /// </summary>
        /// <param name="sender">Event source.</param>
        /// <param name="e">Routed event data.</param>
        /// <remarks>
        /// - Selection is copied before modification to avoid
        ///   collection modification issues.
        /// - Removal is case-insensitive.
        /// - UI is refreshed after updating the queue.
        /// </remarks>
        private void RemoveSelected_Click(object sender, RoutedEventArgs e)
        {
            var selected = FileListBox.SelectedItems.Cast<string>().ToList();
            foreach (var s in selected)
                _files.RemoveAll(x => string.Equals(x, s, StringComparison.OrdinalIgnoreCase));
            
            RefreshFileList();
        }

        /// <summary>
        /// Clears the entire processing queue.
        /// </summary>
        /// <param name="sender">Event source.</param>
        /// <param name="e">Routed event data.</param>
        /// <remarks>
        /// Removes all queued files and refreshes the UI
        /// to reflect the updated state.
        /// </remarks>
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            _files.Clear();
            RefreshFileList();
        }

        /// <summary>
        /// Handles drag-over events and determines whether the dragged
        /// content can be accepted.
        /// </summary>
        /// <param name="sender">Event source.</param>
        /// <param name="e">Drag event data.</param>
        /// <remarks>
        /// - Accepts only file drops.
        /// - Validates that at least one dropped file is a valid .xlsx file.
        /// - Sets appropriate drag effects to control cursor feedback.
        /// </remarks>
        private void Root_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            var dropped = (string[])e.Data.GetData(DataFormats.FileDrop);

            bool hasXlsx = dropped.Any(p =>
                File.Exists(p) &&
                string.Equals(Path.GetExtension(p), ".xlsx", StringComparison.OrdinalIgnoreCase));

            e.Effects = hasXlsx ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        /// <summary>
        /// Handles file drop events and adds valid Excel files
        /// to the processing queue.
        /// </summary>
        /// <param name="sender">Event source.</param>
        /// <param name="e">Drag event data.</param>
        /// <remarks>
        /// - Accepts only file drops.
        /// - Delegates validation and duplicate handling to AddToQueue.
        /// - Marks the event as handled to prevent further propagation.
        /// </remarks>
        private void Root_PreviewDrop(object sender, DragEventArgs e)
        {
            StatusText.Text = "Drop detected";

            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return;

            var dropped = (string[])e.Data.GetData(DataFormats.FileDrop);
            AddToQueue(dropped);

            e.Handled = true;
        }

        /// <summary>
        /// Converts all queued Excel files sequentially into UHD PNG images.
        /// </summary>
        /// <param name="sender">Event source.</param>
        /// <param name="e">Routed event data.</param>
        /// <remarks>
        /// - Runs conversion work on a background thread via <see cref="Task.Run"/> to keep the WPF UI responsive.
        /// - Processes files sequentially because Excel COM interop is not designed for parallel execution.
        /// - Updates progress and status text after each file.
        /// </remarks>
        private async void ConvertAll_Click(object sender, RoutedEventArgs e)
        {
            string outputDir = OutputTextBox.Text?.Trim() ?? "";

            if(string.IsNullOrWhiteSpace(outputDir))
            {
                StatusText.Text = "Please select an output folder.";
                return;
            }
            if (_files.Count == 0)
            {
                StatusText.Text = "Queue is empty. Add .xlsx files first.";
                return;
            }

            Directory.CreateDirectory(outputDir);

            ConvertAllButton.IsEnabled = false;
            ProgressBar.Value = 0;

            try
            {
                var converter = new ExcelToUhdConverter();

                int total = _files.Count;
                for (int i= 0; i < total; i++ )
                {
                    string inputPath = _files[i];
                    string fileName = Path.GetFileName(inputPath);

                    StatusText.Text = $"Converting {i + 1}/{total}: {fileName}";

                    await Task.Run(() =>
                    {
                        converter.Convert(inputPath, outputDir, msg =>
                        {
                            Dispatcher.Invoke(() => StatusText.Text = msg);
                        });
                    });

                    ProgressBar.Value = ((i + 1) * 100.0) / total;
                }

                StatusText.Text = "Done.";
            }
            catch (Exception ex)
            {
                StatusText.Text = "Error: " + ex.Message;
            }
            finally
            {
                ConvertAllButton.IsEnabled = true;
            }
        }
    }
}