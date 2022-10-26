using System;

namespace tff.main.Commands;

internal class StopCommand : BaseCommand
{
    public StopCommand(Action<object> execute, Func<object, bool> canExecute = null) : base(execute, canExecute) { }
}