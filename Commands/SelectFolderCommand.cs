using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using tff.main.Models;

namespace tff.main.Commands
{
    public class SelectFolderCommand  : BaseCommand
    {
        public SelectFolderCommand(Action<object> execute, Func<object, bool> canExecute = null) : base(execute, canExecute)
    {
    }
}
}
