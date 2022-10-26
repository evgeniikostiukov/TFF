using System;

namespace tff.main.Commands;

public class SelectFileCommand : BaseCommand
{
    public SelectFileCommand(Action<object> execute, Func<object, bool> canExecute = null) :
        base(execute, canExecute) { }
}