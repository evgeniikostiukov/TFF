using System.Windows;

namespace tff.main.Models;

public class EntryArgs:DependencyObject
{
    private static readonly DependencyProperty TargetFileProperty;
    private static readonly DependencyProperty EtalonFolderProperty;
    private static readonly DependencyProperty TestFolderProperty;
    private static readonly DependencyProperty SavePathProperty;


    static EntryArgs()
    {
        TargetFileProperty = DependencyProperty.Register("TargetFile", typeof(string), typeof(EntryArgs));
        EtalonFolderProperty = DependencyProperty.Register("EtalonFolder", typeof(string), typeof(EntryArgs));
        TestFolderProperty = DependencyProperty.Register("TestFolder", typeof(string), typeof(EntryArgs));
        SavePathProperty = DependencyProperty.Register("SavePath", typeof(string), typeof(EntryArgs));
    }

    /// <summary>
    /// Файл для обработки
    /// </summary>
    public string TargetFile
    {
        get
        {
            return (string)GetValue(TargetFileProperty);
        }
        set
        {
            SetValue(TargetFileProperty, value);
        }
    }

    /// <summary>
    /// Путь к папке с эталонными запросами и ответами
    /// </summary>
    public string EtalonFolder
    {
        get
        {
            return (string)GetValue(EtalonFolderProperty);
        }
        set
        {
            SetValue(EtalonFolderProperty, value);
        }
    }

    /// <summary>
    /// Путь к папке с тестовыми сценариями
    /// </summary>
    public string TestFolder
    {
        get
        {
            return (string)GetValue(TestFolderProperty);
        }
        set
        {
            SetValue(TestFolderProperty, value);
        }
    }

    /// <summary>
    /// Путь сохранения файла
    /// </summary>
    public string SavePath
    {
        get
        {
            return (string)GetValue(SavePathProperty);
        }
        set
        {
            SetValue(SavePathProperty, value);
        }
    }

    /// <summary>
    /// Результирующий файл
    /// </summary>
    public ResultFile ResultFile { get; set; }
}