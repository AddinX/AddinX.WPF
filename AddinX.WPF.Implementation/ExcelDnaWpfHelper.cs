using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms.Integration;
using AddinX.Wpf.Contract;
using ExcelDna.Integration;

namespace AddinX.Wpf.Implementation
{
    public class ExcelDnaWpfHelper : IWpfHelper
    {
        public async void Show(Window window, TaskCreationOptions taskOption = TaskCreationOptions.None)
        {
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext
                    .SetSynchronizationContext(new ExcelSynchronizationContext());
            }
            var scheduler = TaskScheduler.FromCurrentSynchronizationContext();

            await Task.Factory.StartNew(() =>
            {
                if (null == System.Windows.Application.Current)
                {
                    new System.Windows.Application
                    { ShutdownMode = ShutdownMode.OnExplicitShutdown };
                }
                ElementHost.EnableModelessKeyboardInterop(window);
                if (System.Windows.Application.Current != null)
                {
                    System.Windows.Application.Current.MainWindow = window;
                }
                window.ShowDialog();
            }, CancellationToken.None, taskOption, scheduler);
        }
    }
}
