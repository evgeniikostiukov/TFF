using System.ComponentModel;
using System.Windows;
using System.Runtime.CompilerServices;
namespace tff.main.Models;

public class Entry : INotifyPropertyChanged
{
    private string _targetFile;
    private string _etalonFolder;
    private string _testFolder;
    private string _savePath;
    private int _progress;
    private Visibility _startEnabled;
    private int _totalCount;


    /// <summary>
    /// Файл для обработки
    /// </summary>
    public string TargetFile
    {
        get
        {
            return _targetFile;
        }
        set
        {
            _targetFile = value;
            OnPropertyChanged("TargetFile");
        }
    }

    /// <summary>
    /// Путь к папке с эталонными запросами и ответами
    /// </summary>
    public string EtalonFolder
    {
        get
        {
            return _etalonFolder;
        }
        set
        {
            _etalonFolder = value;
            OnPropertyChanged("EtalonFolder");
        }
    }

    /// <summary>
    /// Путь к папке с тестовыми сценариями
    /// </summary>
    public string TestFolder
    {
        get
        {
            return _testFolder;
        }
        set
        {
            _testFolder = value;
            OnPropertyChanged("TestFolder");
        }
    }

    /// <summary>
    /// Путь сохранения файла
    /// </summary>
    public string SavePath
    {
        get
        {
            return _savePath;
        }
        set
        {
            _savePath = value;
            OnPropertyChanged("SavePath");
        }
    }

    /// <summary>
    /// Прогресс
    /// </summary>
    public int Progress
    {
        get
        {
            return _progress;
        }
        set
        {
            _progress = value;
            OnPropertyChanged("Progress");
        }
    }

    public Visibility StartVisible
    {
        get
        {
            return _startEnabled;
        }
        set
        {
            _startEnabled = value;
            OnPropertyChanged("StartVisible");
            OnPropertyChanged("StopVisible");
        }
    }

    public Visibility StopVisible
    {
        get
        {
            return _startEnabled == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
        }
    }

    public int TotalCount
    {
        get
        {
            return _totalCount;
        }
        set
        {
            _totalCount = value;
            OnPropertyChanged("TotalCount");
        }
    }


    public event PropertyChangedEventHandler? PropertyChanged;
    public void OnPropertyChanged([CallerMemberName] string prop = "")
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }
}