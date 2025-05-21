using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;

namespace GOST_Control
{
    /// <summary>
    /// Сервисный класс для работы с Json
    /// </summary>
    internal class JsonGostService
    {
        private List<Gost> _gosts;
        private readonly string _resourceName;
        private readonly string _externalFilePath;

        /// <summary>
        /// Конструктор с указанием пути к JSON файлу
        /// </summary>
        /// <param name="resourceName">Имя встроенного ресурса</param>
        /// <param name="externalFilePath">Путь к внешнему файлу для сохранения изменений (опционально)</param>
        public JsonGostService(string resourceName, string externalFilePath = null)
        {
            _resourceName = resourceName;
            _externalFilePath = externalFilePath ?? Path.Combine(AppContext.BaseDirectory, "gosts_modified.json");
            _gosts = LoadInitialData();
        }

        /// <summary>
        /// Загрузка действубщего
        /// </summary>
        /// <returns></returns>
        private List<Gost> LoadInitialData()
        {
            // Сначала пытаемся загрузить из внешнего файла (если есть изменения)
            if (File.Exists(_externalFilePath))
            {
                try
                {
                    var json = File.ReadAllText(_externalFilePath);
                    return JsonSerializer.Deserialize<List<Gost>>(json) ?? LoadFromEmbeddedResource();
                }
                catch
                {
                    return LoadFromEmbeddedResource();
                }
            }

            return LoadFromEmbeddedResource();
        }

        /// <summary>
        /// Загрузка вшитого файла из папки "Resources"
        /// </summary>
        /// <returns></returns>
        /// <exception cref="FileNotFoundException"></exception>
        private List<Gost> LoadFromEmbeddedResource()
        {
            var assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(_resourceName))
            {
                if (stream == null)
                    throw new FileNotFoundException($"Resource '{_resourceName}' not found.");

                using (StreamReader reader = new StreamReader(stream))
                {
                    string json = reader.ReadToEnd();
                    return JsonSerializer.Deserialize<List<Gost>>(json) ?? new List<Gost>();
                }
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
                await File.WriteAllTextAsync(_externalFilePath, json);
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка сохранения данных в JSON: {ex.Message}");
            }
        }

        /// <summary>
        /// Получает ГОСТ по его идентификатору
        /// </summary>
        public async Task<Gost?> GetGostByIdAsync(int gostId)
        {
            return await Task.FromResult(_gosts.Find(g => g.GostId == gostId));
        }

        /// <summary>
        /// Получает все ГОСТы
        /// </summary>
        public async Task<List<Gost>> GetAllGostsAsync()
        {
            return await Task.FromResult(_gosts);
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
    }
}