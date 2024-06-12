using System.Windows;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace RevitAddin
{
    public partial class FolderSelectionWindow : Window
    {
        public string SelectedFolderPath { get; private set; }

        public FolderSelectionWindow()
        {
            InitializeComponent();
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    SelectedFolderPath = dialog.FileName;
                    FolderPathTextBox.Text = SelectedFolderPath;
                }
            }
        }

        private void AttachButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedFolderPath))
            {
                MessageBox.Show("Please select a folder first.");
                return;
            }

            DialogResult = true;
            Close();
        }
    }
}
    