using System.Windows.Input;

namespace Task.Parser.XLSX.ViewModels;

/// <summary>
/// Главная модель представления: пока только строка состояния и команда
/// загрузки файла; сама загрузка появится позже.
/// </summary>
public sealed class MainViewModel : ViewModelBase
{
    private string _statusText = "Файл не загружен. Нажмите «Загрузить файл…»";

    public string StatusText
    {
        get => _statusText;
        private set { _statusText = value; OnPropertyChanged(); }
    }

    public ICommand LoadFileCommand { get; }

    public MainViewModel()
    {
        LoadFileCommand = new RelayCommand(_ => StatusText = "Загрузка файла…");
    }
}
