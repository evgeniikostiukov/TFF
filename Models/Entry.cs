using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace tff.main.Models;

public class Entry : INotifyPropertyChanged
{
    private string _currentTemplate;
    private string _etalonFolder;
    private int _progress;
    private string _savePath;
    private Visibility _startEnabled;
    private string _targetFile;
    private string _testFolder;
    private int _totalCount;

    /// <summary>
    ///     Файл для обработки
    /// </summary>
    public string TargetFile
    {
        get => _targetFile;
        set
        {
            _targetFile = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    ///     Путь к папке с эталонными запросами и ответами
    /// </summary>
    public string EtalonFolder
    {
        get => _etalonFolder;
        set
        {
            _etalonFolder = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    ///     Путь к папке с тестовыми сценариями
    /// </summary>
    public string TestFolder
    {
        get => _testFolder;
        set
        {
            _testFolder = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    ///     Путь сохранения файла
    /// </summary>
    public string SavePath
    {
        get => _savePath;
        set
        {
            _savePath = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    ///     Прогресс
    /// </summary>
    public int Progress
    {
        get => _progress;
        set
        {
            _progress = value;
            OnPropertyChanged();
        }
    }

    public string CurrentTemplate
    {
        get => _currentTemplate;
        set
        {
            _currentTemplate = value;
            OnPropertyChanged();
        }
    }

    public Visibility StartVisible
    {
        get => _startEnabled;
        set
        {
            _startEnabled = value;
            OnPropertyChanged();
            OnPropertyChanged("StopVisible");
        }
    }

    public Visibility StopVisible => _startEnabled == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;

    public int TotalCount
    {
        get => _totalCount;
        set
        {
            _totalCount = value;
            OnPropertyChanged();
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public void OnPropertyChanged([CallerMemberName] string prop = "")
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }
}