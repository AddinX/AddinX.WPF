using System.Windows;
using System.Windows.Forms.Integration;
using AddinX.Wpf.Contract;
using ExcelDna.Integration;

namespace AddinX.Wpf.Implementation
{
    public class ExcelDnaWpfHelper : IWpfHelper
    {
        public void Show(Window window)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
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
            });
        }
    }
}
