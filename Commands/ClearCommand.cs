using System;

namespace tff.main.Commands
{
    public class ClearCommand : BaseCommand
    {
        public ClearCommand(Action<object> execute, Func<object, bool> canExecute = null) : base(execute, canExecute)
        {
        }
    }
}
