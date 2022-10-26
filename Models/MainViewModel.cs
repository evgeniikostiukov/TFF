using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using tff.main.Commands;
using tff.main.Handlers;
using MessageBox = System.Windows.Forms.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace tff.main.Models;

internal class MainViewModel : INotifyPropertyChanged
{
    private SelectFileCommand _selectFileCommand;

    private SelectFolderCommand _selectFolderCommand;

    private StartCommand _startCommand;

    private StopCommand _stopCommand;

    public MainViewModel()
    {
        EntryEntity = new Entry
        {
            StartVisible = Visibility.Visible,
        };
    }

    public Entry EntryEntity { get; set; }

    public SelectFolderCommand SelectFolderCommand
    {
        get
        {
            return _selectFolderCommand ??= new SelectFolderCommand(obj =>
                       {
                           var dialog = new FolderBrowserDialog
                           {
                               Description = "Выберите папку",
                               UseDescriptionForTitle = true,
                               ShowNewFolderButton = true,
                           };

                           var result = dialog.ShowDialog();

                           if (result == DialogResult.OK)
                           {
                               SetPropertyValue(obj, dialog.SelectedPath);
                           }
                       }
                   );
        }
    }

    public SelectFileCommand SelectFileCommand
    {
        get
        {
            return _selectFileCommand ??= new SelectFileCommand(obj =>
                       {
                           var dialog = new OpenFileDialog
                           {
                               AddExtension = true,
                               CheckFileExists = true,
                               Multiselect = false,
                               Filter = "MS Word documents|*.docx;*.doc",
                           };

                           var result = dialog.ShowDialog();

                           if (result == true)
                           {
                               SetPropertyValue(obj, dialog.FileName);
                           }
                       }
                   );
        }
    }

    public StartCommand StartCommand
    {
        get
        {
            return _startCommand ??= new StartCommand(_ =>
                       {
                           foreach (var propertyInfo in EntryEntity.GetType()
                                        .GetProperties())
                               if (propertyInfo.PropertyType == typeof(string)
                                   && propertyInfo.GetValue(EntryEntity) == null)
                               {
                                   MessageBox.Show("Заполните все поля", "Ошибка");

                                   return;
                               }

                           DocProcessHandler.Execute(EntryEntity);
                           MessageBox.Show("Задача запущена");
                       }
                   );
        }
    }

    public StopCommand StopCommand
    {
        get { return _stopCommand ??= new StopCommand(_ => { DocProcessHandler.worker_Stop(); }); }
    }

    public event PropertyChangedEventHandler PropertyChanged;

    public void OnPropertyChanged([CallerMemberName] string prop = "")
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }

    private PropertyInfo? GetProperty(string propertyName)
    {
        return typeof(Entry).GetProperty(propertyName);
    }

    private static T? ConvertArgument<T>(object obj)
        where T : class
    {
        return obj as T;
    }

    private void SetPropertyValue(object argument, string value)
    {
        var propertyName = ConvertArgument<string>(argument);

        if (propertyName == null)
        {
            return;
        }

        var property = GetProperty(propertyName);
        property?.SetValue(EntryEntity, value);
    }
}