using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace GOST_Control;

public partial class SettingGost : Window
{
    public SettingGost()
    {
        InitializeComponent();
    }

    /// <summary>
    /// ����� �� ���� ��������
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_Setting_Out(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        Close();
    }

    /// <summary>
    /// ���������� ��������� � �����
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_Save(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {






    }
}