using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;

namespace GOST_Control
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // ������� ���� gosts.json ��� ������ �������, ���� ��� ���
            string jsonPath = Path.Combine(AppContext.BaseDirectory, "gosts.json");
            if (!File.Exists(jsonPath))
            {
                File.WriteAllText(jsonPath, "[]");
            }
        }

        private readonly FileDialogFilter fileFilter = new()
        {
            Extensions = new List<string>() { "doc", "docx" },
            Name = "��������� Word"
        };

        private async void Button_Click_SelectFile(object? sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new OpenFileDialog
                {
                    Title = "�������� �������� ��� ��������",
                    AllowMultiple = false,
                    Filters = { fileFilter }
                };

                var result = await dialog.ShowAsync(this);

                if (result != null && result.Length > 0)
                {
                    string filePath = result[0];

                    GOST_�heck checkWindow = new GOST_�heck(filePath);
                    checkWindow.Show();
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"������: {ex.Message}");
            }
        }

        /// <summary>
        /// ��������� �����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Button_Click_Setting(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            this.Classes.Add("blur-effect");

            SettingGost settingGost = new SettingGost();
            await settingGost.ShowDialog(this);

            this.Classes.Remove("blur-effect");
        }
    }
}