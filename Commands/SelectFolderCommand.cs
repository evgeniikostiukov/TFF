using System;

namespace tff.main.Commands;

public class SelectFolderCommand : BaseCommand
{
    public SelectFolderCommand(Action<object> execute, Func<object, bool> canExecute = null) :
        base(execute, canExecute) { }
}