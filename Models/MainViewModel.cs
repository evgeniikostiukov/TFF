using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using tff.main.Commands;
using tff.main.Extensions;
using tff.main.Handlers;
using MessageBox = System.Windows.Forms.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace tff.main.Models;

internal class MainViewModel : INotifyPropertyChanged
{
    private readonly DocProcessHandler _docProcessHandler;

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

        _docProcessHandler = new DocProcessHandler(EntryEntity);
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
                    var filter = "MS Word documents|*.docx;*.doc";
                    var command = obj;

                    if (obj != null && obj.ToString()!.StartsWith('['))
                    {
                        var paramArray = (obj as string)?.Trim('[', ']').Split(',');

                        if (paramArray is null)
                        {
                            throw new InvalidOperationException("Не переданы параметры фильтра файлов");
                        }

                        command = paramArray[0];
                        filter = paramArray[1];
                    }

                    var dialog = new OpenFileDialog
                    {
                        AddExtension = true,
                        CheckFileExists = true,
                        Multiselect = false,
                        Filter = filter,
                    };

                    var result = dialog.ShowDialog();

                    if (result == true)
                    {
                        SetPropertyValue(command, dialog.FileName);
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
                    var isWordOrSaveEmptyMsg =
                        $"\"{EntryEntity.GetDisplayName(nameof(Entry.TargetFile))}\", \"{EntryEntity.GetDisplayName(nameof(Entry.SavePath))}\"";

                    var isXsdOrFoldersEmptyMsg =
                        $"\"{EntryEntity.GetDisplayName(nameof(EntryEntity.TargetXsdFile))}\", \"{EntryEntity.GetDisplayName(nameof(EntryEntity.EtalonFolder))}\" или \"{EntryEntity.GetDisplayName(nameof(EntryEntity.TestFolder))}\"";

                    var isWordOrSaveEmpty = EntryEntity.TargetFile == null || EntryEntity.SavePath == null;

                    var isXsdOrFoldersEmpty = EntryEntity.EtalonFolder == null
                     && EntryEntity.TestFolder == null
                     && EntryEntity.TargetXsdFile == null;

                    if (isWordOrSaveEmpty || isXsdOrFoldersEmpty)
                    {
                        var result = isWordOrSaveEmpty ? isWordOrSaveEmptyMsg : isXsdOrFoldersEmptyMsg;
                        MessageBox.Show($"Заполните следующие поля: {result}", "Ошибка");

                        return;
                    }

                    _docProcessHandler.Execute();
                }
            );
        }
    }

    public StopCommand StopCommand
    {
        get { return _stopCommand ??= new StopCommand(_ => { _docProcessHandler.worker_Stop(); }); }
    }

    public event PropertyChangedEventHandler PropertyChanged;

    public void OnPropertyChanged([CallerMemberName] string prop = "")
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }

    private PropertyInfo GetProperty(string propertyName)
    {
        return typeof(Entry).GetProperty(propertyName);
    }

    private static T ConvertArgument<T>(object obj)
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