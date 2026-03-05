using System.Windows;
using BillMatch.Wpf.ViewModels;

namespace BillMatch.Wpf;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainViewModel();
    }
}
