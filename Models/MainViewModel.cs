using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using tff.main.Commands;
using tff.main.Handlers;
using MessageBox = System.Windows.Forms.MessageBox;

namespace tff.main.Models
{
    internal class MainViewModel : INotifyPropertyChanged
    {
        public Entry EntryEntity { get; set; }

        public MainViewModel()
        {
            EntryEntity = new Entry() { StartVisible = Visibility.Visible };
        }

        private SelectFolderCommand? _selectFolderCommand;

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

        private SelectFileCommand? _selectFileCommand;

        public SelectFileCommand SelectFileCommand
        {
            get
            {
                return _selectFileCommand ??= new SelectFileCommand(obj =>
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
                                   SetPropertyValue(obj, dialog.FileName);
                               }
                           }
                       );
            }
        }

        private StartCommand? _startCommand;

        public StartCommand StartCommand
        {
            get
            {
                return _startCommand ??= new StartCommand(_ =>
                   {
                       foreach (var propertyInfo in EntryEntity.GetType().GetProperties())
                       {
                           if (propertyInfo.PropertyType == typeof(string)
                               && propertyInfo.GetValue(EntryEntity) == null)
                           {
                               MessageBox.Show("Заполните все поля", "Ошибка");
                               return;
                           }
                       }

                       DocProcessHandler.Execute(EntryEntity);
                       MessageBox.Show("Задача запущена");
                   }
               );
            }
        }

        private StopCommand _stopCommand;

        public StopCommand StopCommand
        {
            get
            {
                return _stopCommand ??= new StopCommand(_ =>
                  {
                   
                          DocProcessHandler.worker_Stop();
                   
                  });
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }



        private PropertyInfo? GetProperty(string propertyName)
        {
            return typeof(Entry).GetProperty(propertyName);
        }

        private static T? ConvertArgument<T>(object obj)
            where T : class =>
            obj as T;

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
}