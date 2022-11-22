using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
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
    private string _targetXsdFile;
    private string _testFolder;
    private int _totalCount;

    /// <summary>
    ///     Файл для обработки
    /// </summary>
    [Display(Name = "Файл для обработки")]
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
    ///     XSD схема
    /// </summary>
    [Display(Name = "XSD схема")]
    public string TargetXsdFile
    {
        get => _targetXsdFile;
        set
        {
            _targetXsdFile = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    ///     Путь к папке с эталонными запросами и ответами
    /// </summary>
    [Display(Name = "Путь к папке с эталонными запросами и ответами")]
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
    [Display(Name = "Путь к папке с тестовыми сценариями")]
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
    [Display(Name = "Путь сохранения файла")]
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
    [Display(Name = "Прогресс")]
    public int Progress
    {
        get => _progress;
        set
        {
            _progress = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    ///     Текущий шаблон
    /// </summary>
    [Display(Name = "Текущий шаблон")]
    public string CurrentTemplate
    {
        get => _currentTemplate;
        set
        {
            _currentTemplate = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    ///     Видимость кнопки Начать
    /// </summary>
    [Display(Name = "Видимость кнопки Начать")]
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

    /// <summary>
    ///     Видимость кнопки Остановить
    /// </summary>
    [Display(Name = "Видимость кнопки Остановить")]
    public Visibility StopVisible => _startEnabled == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;

    /// <summary>
    ///     Всего шаблонов
    /// </summary>
    [Display(Name = "Всего шаблонов")]
    public int TotalCount
    {
        get => _totalCount;
        set
        {
            _totalCount = value;
            OnPropertyChanged();
        }
    }

    public event PropertyChangedEventHandler PropertyChanged;

    public void OnPropertyChanged([CallerMemberName] string prop = "")
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }
}