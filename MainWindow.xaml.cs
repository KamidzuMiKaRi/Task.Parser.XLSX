using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
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

        _viewModel.ErrorOccurred += (_, message) =>
            MessageBox.Show(this, message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        _viewModel.PropertyChanged += OnViewModelPropertyChanged;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        // При новой загрузке файла пересоздаём колонки под заголовки.
        if (e.PropertyName == nameof(MainViewModel.Headers))
            BuildColumns();
    }

    private void BuildColumns()
    {
        Grid.Columns.Clear();
        for (int i = 0; i < _viewModel.Headers.Count; i++)
        {
            Grid.Columns.Add(new DataGridTextColumn
            {
                Header = _viewModel.Headers[i],
                Binding = new Binding($"[{i}]")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                },
                Width = new DataGridLength(130),
            });
        }
    }
}
