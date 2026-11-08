using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.UI.Xaml.Shapes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.ApplicationModel;
using Windows.ApplicationModel.Activation;
using Windows.Foundation;
using Windows.Foundation.Collections;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace MailMind
{
    /// <summary>
    /// Provides application-specific behavior to supplement the default Application class.
    /// </summary>
    public partial class App : Application
    {
        private Window? _window;

        /// <summary>
        /// Initializes the singleton application object.  This is the first line of authored code
        /// executed, and as such is the logical equivalent of main() or WinMain().
        /// </summary>
        public App()
        {
            Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride = Strings.LanguageTag;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            // MIME headers without a declared charset are common in older Chinese mail.
            MimeKit.ParserOptions.Default.CharsetEncoding = System.Text.Encoding.GetEncoding("GB18030");
            InitializeComponent();
        }

        /// <summary>
        /// Invoked when the application is launched.
        /// </summary>
        /// <param name="args">Details about the launch request and process.</param>
        protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            var window = new MainWindow();
            _window = window;
            _window.Activate();
            Program.ActivationPending = () => window.DispatcherQueue.TryEnqueue(ProcessPendingActivations);
            HandleActivation(Microsoft.Windows.AppLifecycle.AppInstance.GetCurrent().GetActivatedEventArgs());
            ProcessPendingActivations();
        }

        private void ProcessPendingActivations()
        {
            while (Program.TryDequeue(out var activation))
                if (activation != null) HandleActivation(activation);
        }

        private void HandleActivation(Microsoft.Windows.AppLifecycle.AppActivationArguments activation)
        {
            var window = (MainWindow)_window!;
            window.AppWindow.Show();
            window.Activate();
            if (activation.Kind == Microsoft.Windows.AppLifecycle.ExtendedActivationKind.Protocol &&
                activation.Data is IProtocolActivatedEventArgs protocol && protocol.Uri.Scheme.Equals("mailto", StringComparison.OrdinalIgnoreCase))
            {
                window.DispatcherQueue.TryEnqueue(async () => await window.OpenMailtoAsync(protocol.Uri));
            }
        }
    }
}
