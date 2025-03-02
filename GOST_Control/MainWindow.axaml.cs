using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Avalonia.Controls;

namespace GOST_Control
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private readonly FileDialogFilter fileFilter = new()
        {
            Extensions = new List<string>() { "doc", "docx" }, // Допустимые форматы
            Name = "Документы Word"
        };

        private async void Button_Click_SelectFile(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            try
            {
                // Создаем диалог выбора файла
                OpenFileDialog dialog = new OpenFileDialog
                {
                    Title = "Выберите документ для проверки",
                    AllowMultiple = false
                };
                // Добавляем фильтр
                dialog.Filters.Add(fileFilter);
                // Открываем диалог
                var result = await dialog.ShowAsync(this);
                // Проверяем, выбран ли файл
                if (result is { Length: > 0 })
                {
                    string filePath = result[0];

                    // Открываем новое окно
                    var checkWindow = new GOST_Сheck(filePath);
                    checkWindow.Show();
                    this.Close(); // Закрываем главное окно

                    // обрабатываем файл
                   // await ProcessSelectedFileAsync(filePath);
                }
            }
            catch { }
        }

        private void Button_Click_Setting(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            Settings settings = new Settings();
            settings.Show();
            Close();
        }   
    }
}