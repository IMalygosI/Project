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
        /*
        private async Task ProcessSelectedFileAsync(string filePath)
        {
            try
            {
                // �������� ������������� �����
                if (!File.Exists(filePath))
                    throw new FileNotFoundException("���� �� ������", filePath);
                // �������� ��������� ����� (������ ����� �������� �������� ������)
                string fileName = Path.GetFileName(filePath);
                long fileSize = new FileInfo(filePath).Length;
                // �������� ���� � ����� ������� (�� �������� � �������������)
                string destinationPath = Path.Combine("ProcessedDocs", fileName);
                // ��������, ��� ����� ����������
                Directory.CreateDirectory("ProcessedDocs");
                // �������� ����
                await Task.Run(() => File.Copy(filePath, destinationPath, overwrite: true));
            }
            catch { }
        }
        */


    }
}