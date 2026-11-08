using System;
using System.Collections.Concurrent;
using Microsoft.Windows.AppLifecycle;

namespace MailMind;

internal static class Program
{
    private static readonly ConcurrentQueue<AppActivationArguments> Pending = new();
    internal static Action? ActivationPending;

    [STAThread]
    private static void Main(string[] args)
    {
        WinRT.ComWrappersSupport.InitializeComWrappers();
        var activation = AppInstance.GetCurrent().GetActivatedEventArgs();
        var instance = AppInstance.FindOrRegisterForKey("MailMind");
        if (!instance.IsCurrent)
        {
            System.Threading.Tasks.Task.Run(async () => await instance.RedirectActivationToAsync(activation)).GetAwaiter().GetResult();
            return;
        }

        instance.Activated += (_, eventArgs) =>
        {
            Pending.Enqueue(eventArgs);
            ActivationPending?.Invoke();
        };
        Microsoft.UI.Xaml.Application.Start(callbackArgs =>
        {
            var context = new Microsoft.UI.Dispatching.DispatcherQueueSynchronizationContext(Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread());
            System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            _ = new App();
        });
    }

    internal static bool TryDequeue(out AppActivationArguments? activation) => Pending.TryDequeue(out activation);
}
