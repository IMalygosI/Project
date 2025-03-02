using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media.TextFormatting;
using DocumentFormat.OpenXml.Vml;
using GOST_Control.Models;

namespace GOST_Control;

public partial class Settings : Window
{
    List<Gost> gosts = new List<Gost>();
    public Settings()
    {
        InitializeComponent();

    }    

    /// <summary>
    /// Сохранение настроек ГОСТа и возврат в главное меню
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_Save(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {








        MainWindow mainWindow = new MainWindow();
        mainWindow.Show();
        Close();
    }

    /// <summary>
    /// Выход в главное меню
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_Logout(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        MainWindow mainWindow = new MainWindow();
        mainWindow.Show();
        Close();
    }
}