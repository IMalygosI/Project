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
            Extensions = new List<string>() { "doc", "docx" }, // ���������� �������
            Name = "��������� Word"
        };

        private async void Button_Click_SelectFile(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            try
            {
                // ������� ������ ������ �����
                OpenFileDialog dialog = new OpenFileDialog
                {
                    Title = "�������� �������� ��� ��������",
                    AllowMultiple = false
                };
                // ��������� ������
                dialog.Filters.Add(fileFilter);
                // ��������� ������
                var result = await dialog.ShowAsync(this);
                // ���������, ������ �� ����
                if (result is { Length: > 0 })
                {
                    string filePath = result[0];

                    // ��������� ����� ����
                    var checkWindow = new GOST_�heck(filePath);
                    checkWindow.Show();
                    this.Close(); // ��������� ������� ����

                    // ������������ ����
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