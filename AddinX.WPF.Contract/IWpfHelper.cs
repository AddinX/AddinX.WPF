using System.Threading.Tasks;
using System.Windows;

namespace AddinX.Wpf.Contract
{
    public interface IWpfHelper
    {
        void Show(Window window, TaskCreationOptions taskOption = TaskCreationOptions.None);
    }
}
