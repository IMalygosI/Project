using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace GOST_Control
{
    /// <summary>
    /// Класс работы с Json
    /// </summary>
    internal class JsonGostService
    {
        private readonly string _jsonFilePath;
        private List<Gost> _gosts;

        /// <summary>
        /// Конструктор с указанием пути к JSON файлу
        /// </summary>
        /// <param name="jsonFilePath">Путь к файлу с данными ГОСТов</param>
        public JsonGostService(string jsonFilePath)
        {
            _jsonFilePath = jsonFilePath;
            _gosts = new List<Gost>();
            LoadGosts().Wait(); // Загружаем данные при создании сервиса
        }

        /// <summary>
        /// Загружает данные ГОСТов из JSON файла
        /// </summary>
        private async Task LoadGosts()
        {
            try
            {
                if (!File.Exists(_jsonFilePath))
                {
                    // Создаем пустой файл, если его нет
                    await SaveGosts();
                    return;
                }

                var json = await File.ReadAllTextAsync(_jsonFilePath);
                _gosts = JsonSerializer.Deserialize<List<Gost>>(json) ?? new List<Gost>();
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка загрузки данных из JSON: {ex.Message}");
            }
        }

        /// <summary>
        /// Сохраняет данные ГОСТов в JSON файл
        /// </summary>
        private async Task SaveGosts()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(_gosts, options);
                await File.WriteAllTextAsync(_jsonFilePath, json);
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка сохранения данных в JSON: {ex.Message}");
            }
        }

        /// <summary>
        /// Получает ГОСТ по его идентификатору
        /// </summary>
        /// <param name="gostId">Идентификатор ГОСТа</param>
        /// <returns>Найденный ГОСТ или null</returns>
        public async Task<Gost?> GetGostByIdAsync(int gostId)
        {
            return _gosts.Find(g => g.GostId == gostId);
        }

        /// <summary>
        /// Получает все ГОСТы
        /// </summary>
        public async Task<List<Gost>> GetAllGostsAsync()
        {
            return _gosts;
        }

        /// <summary>
        /// Добавляет или обновляет ГОСТ
        /// </summary>
        public async Task AddOrUpdateGostAsync(Gost gost)
        {
            var existing = _gosts.FindIndex(g => g.GostId == gost.GostId);
            if (existing >= 0)
            {
                _gosts[existing] = gost;
            }
            else
            {
                _gosts.Add(gost);
            }
            await SaveGosts();
        }

        /// <summary>
        /// Удаляет ГОСТ
        /// </summary>
        public async Task DeleteGostAsync(int gostId)
        {
            _gosts.RemoveAll(g => g.GostId == gostId);
            await SaveGosts();
        }
    }
}
