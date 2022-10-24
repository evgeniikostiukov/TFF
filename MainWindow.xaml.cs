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
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using tff.main.Models;

namespace tff.main;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private readonly EntryArgs entryArgs;

    public MainWindow()
    {
        InitializeComponent();

        entryArgs = (EntryArgs)Resources["EntryArgs"];

        SelectTargetFileBtn.Click += SelectFile;
        SelectEtalonFolderBtn.Click += SelectFolder;
    }

    private void SelectFolder(object sender, RoutedEventArgs e)
    {
        var dialog = new FolderBrowserDialog
        {
            Description = "Выберите папку",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true,
        };

        var result = dialog.ShowDialog();

        if (result == System.Windows.Forms.DialogResult.OK)
        {
            entryArgs.EtalonFolder = dialog.SelectedPath;
        }
    }

    private void SelectFile(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            AddExtension = true,
            CheckFileExists = true,
            Multiselect = false,
            Filter = "MS Word documents|*.docx;*.doc",
        };

        var result = dialog.ShowDialog();

        if (result == true)
        {
            entryArgs.TargetFile = dialog.FileName;
        }

    }
}