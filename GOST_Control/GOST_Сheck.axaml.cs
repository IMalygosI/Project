using System;
using System.IO;
using System.Linq;  // Для методов расширений LINQ
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GOST_Control.Context;
using GOST_Control.Models;
using Microsoft.EntityFrameworkCore;

namespace GOST_Control;

public partial class GOST_Сheck : Window
{
    public DbSet<Gost> Gosts { get; set; }
    private readonly string _filePath;

    public GOST_Сheck(string filePath)
    {
        InitializeComponent();
        _filePath = filePath; // Сохраняем путь к файлу
        FilePathTextBlock.Text = $"Путь к файлу: {_filePath}"; // Отображаем путь в UI
    }

    /// <summary>
    /// Метод отвечающий за поиск ГОСТа в базе
    /// </summary>
    /// <param name="gostId"></param>
    /// <returns></returns>
    private async Task<Gost> GetGostByIdAsync(int gostId)
    {
        using (var context = new DimaBaseContext())
        {
            return await context.Gosts
                .FirstOrDefaultAsync(g => g.GostId == gostId);  // Просто ищем ГОСТ по его ID
        }
    }














    /// <summary>
    /// Метод проверки на ГОСТ
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="gostId"></param>
    /// <returns></returns>
    public async Task CheckFileForGostAsync(string filePath, int gostId)
    {
        // Получаем ГОСТ из базы данных
        var gost = await GetGostByIdAsync(gostId);
        if (gost == null)
        {
            ErrorControlGost.Text = "ГОСТ не найден в базе данных.";
            return;
        }
        else
        {
            ErrorControlGost.Text = "ГОСТ найден в базе данных.";
        }
        
        if (!string.IsNullOrEmpty(gost.FontName))
        {
            try
            {
                using (var wordDoc = WordprocessingDocument.Open(filePath, false))
                {
                    // Обновляем UI
                    if (wordDoc != null)
                    {
                        ErrorControl.Text = "Удалось открыть документ.";
                    }
                    else
                    {
                        ErrorControl.Text = "Не удалось открыть документ.";
                    }

                    ErrorControlFontBD.Text = "7000000";

                    // Забираем текст из документа
                    var body = wordDoc.MainDocumentPart.Document.Body;

                    // Проверки ГОСТа
                    // Проверка шрифта на соответствие госту
                    bool isValid = CheckFont(gost.FontName, body); // Проверка шрифта







                    if (isValid)
                    {
                        GostControl.Text = "Документ соответствует ГОСТу (по шрифту и размеру).";
                    }
                    else
                    {
                        GostControl.Text = "Документ не соответствует ГОСТу (по шрифту и размеру).";
                    }
                }
            }
            catch (Exception ex)
            {
                GostControl.Text = $"Ошибка при открытии документа: {ex.Message}";
            }
        }
        else
        {
            ErrorControlFontBD.Text = "Шрифт или размер шрифта не указаны для этого ГОСТа.";
        }
    }

    /// <summary>
    /// Метод проверки шрифта в документе (тип и размер)
    /// </summary>
    /// <param name="requiredFontName"></param>
    /// <param name="body"></param>
    /// <returns></returns>
    private bool CheckFont(string requiredFontName, Body body)
    {
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            foreach (var run in paragraph.Elements<Run>())
            {
                var runProperties = run.RunProperties;
                if (runProperties == null) continue;

                // Получаем имя шрифта
                var fontNameElement = runProperties.RunFonts;
                string fontName = fontNameElement?.Ascii?.Value;

                if (fontName != null && fontName != requiredFontName)
                {
                    return false; // Если хоть один абзац с другим шрифтом — не соответствует ГОСТу
                }
            }
        }

        return true; // Всё соответствует ГОСТу
    }






























    /// <summary>
    /// Кнопка проверки на соответствие ГОСТу
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private async void Button_Click_SelectFile(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            // Указываем номер ГОСТа
            int gostId = 1; 

            // Проверяем файл на соответствие ГОСТу
            await CheckFileForGostAsync(_filePath, gostId);
        }
        catch (Exception ex)
        {
            GostControl.Text = $"Ошибка: {ex.Message}";
        }
    }
    /// <summary>
    /// Кнопка выхода
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_LogOut(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        MainWindow mainWindow = new MainWindow();
        mainWindow.Show();
        Close();
    }
}
