using System.Windows;
using Task.Parser.XLSX.ViewModels;

namespace Task.Parser.XLSX;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();

        _viewModel = new MainViewModel();
        DataContext = _viewModel;
    }
}
