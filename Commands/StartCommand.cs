using System;

namespace tff.main.Commands;

public class StartCommand : BaseCommand
{
    public StartCommand(Action<object> execute, Func<object, bool> canExecute = null) : base(execute, canExecute) { }
}