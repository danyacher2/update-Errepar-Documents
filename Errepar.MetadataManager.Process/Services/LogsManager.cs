using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Errepar.MetadataManager.Process.Services
    {
    /// <summary>
    /// Gestor de logs para el proceso de Metadata Manager.
    /// Permite escribir, leer y ac
    /// </summary>tualizar logs tanto en archivos locales como en SharePoint.
    public class ItemLogEjecucion
        {
        public int ItemId { get; set; }
        public DateTime Fecha { get; set; }
        public string Mensaje { get; set; }
        public string Estado { get; set; }
        }

    public class ItemLogError
        {
        public string ItemId { get; set; }
        public DateTime Fecha { get; set; }
        public string Mensaje { get; set; }
        public string Estado { get; set; }
        }


    public class ItemLogActivosProcesados
        {
        public int ItemCambiosId { get; set; }
        public int ItemListId { get; set; }
        public int TimeStamp { get; set; }
        public string UrlItem { get; set; }
        public string LibraryName { get; set; }
        public string Activo { get; set; }
        public bool Procesado { get; set; }
        public DateTime Fecha { get; set; }
        }

    public class LogsManager
        {
        private readonly HttpClient _http;
        private readonly string _siteUrl;
        private readonly string _listTitle;

        public LogsManager(HttpClient httpClient, string siteUrl, string listTitle = "Metadata Manager")
            {
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _siteUrl = siteUrl?.TrimEnd('/') ?? throw new ArgumentNullException(nameof(siteUrl));
            _listTitle = listTitle;
            }

        private readonly List<ItemLogEjecucion> _logEjecucionMemoria = new();
        private readonly List<ItemLogActivosProcesados> _logActivosMemoria = new();
        private readonly List<ItemLogError> _logErroresMemoria = new();

        private readonly string _baseLogsPath =
            Path.Combine(
                AppContext.BaseDirectory,
            "Logs"
            );

        private string GetItemFolder(int itemId)
            {
            var path = Path.Combine(_baseLogsPath, $"Item_{itemId}");
            Directory.CreateDirectory(path);
            return path;
            }

        public async Task SaveLogEjecucion(int itemId, ItemLogEjecucion log)
            {
            var folder = GetItemFolder(itemId);
            var file = Path.Combine(folder, "LogEjecucion.json");

            var logs = new List<ItemLogEjecucion>();

            if (File.Exists(file))
                {
                var json = await File.ReadAllTextAsync(file);
                logs = JsonSerializer.Deserialize<List<ItemLogEjecucion>>(json) ?? new();
                }

            logs.Add(log);

            var newJson = JsonSerializer.Serialize(logs, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(file, newJson);
            }

        public async Task SaveLogError(int itemId, ItemLogError log)
            {
            var folder = GetItemFolder(itemId);
            var file = Path.Combine(folder, "LogError.json");

            var logs = new List<ItemLogError>();

            if (File.Exists(file))
                {
                var json = await File.ReadAllTextAsync(file);
                logs = JsonSerializer.Deserialize<List<ItemLogError>>(json) ?? new();
                }

            logs.Add(log);

            var newJson = JsonSerializer.Serialize(logs, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(file, newJson);
            }
        public async Task SaveLogErrorCambios(int itemId, ItemLogError log)
            {
            var folder = GetItemFolder(itemId);
            var file = Path.Combine(folder, "LogErrorCambios.json");

            var logs = new List<ItemLogError>();

            if (File.Exists(file))
                {
                var json = await File.ReadAllTextAsync(file);
                logs = JsonSerializer.Deserialize<List<ItemLogError>>(json) ?? new();
                }

            logs.Add(log);

            var newJson = JsonSerializer.Serialize(logs, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(file, newJson);
            }


        public async Task SaveLogActivosProcesados(int itemId, IEnumerable<ItemLogActivosProcesados> activos)
            {
            var folder = GetItemFolder(itemId);
            var file = Path.Combine(folder, "ActivosProcesados.json");

            var logs = new List<ItemLogActivosProcesados>();

            if (File.Exists(file))
                {
                var json = await File.ReadAllTextAsync(file);
                logs = JsonSerializer.Deserialize<List<ItemLogActivosProcesados>>(json) ?? new();
                }

            foreach (var activo in activos)
                logs.Add(activo);

            var newJson = JsonSerializer.Serialize(logs, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(file, newJson);
            }

        public async Task SyncLogs(int itemId, CancellationToken ct)
            {
            var folder = GetItemFolder(itemId);

            var ejecucionPath = Path.Combine(folder, "LogEjecucion.json");
            var activosPath = Path.Combine(folder, "ActivosProcesados.json");
            var erroresPath = Path.Combine(folder, "LogError.json");

            var ejecucionJson = JsonSerializer.Serialize(_logEjecucionMemoria, new JsonSerializerOptions { WriteIndented = true });
            var activosJson = JsonSerializer.Serialize(_logActivosMemoria, new JsonSerializerOptions { WriteIndented = true });
            var errorJson = JsonSerializer.Serialize(_logErroresMemoria, new JsonSerializerOptions { WriteIndented = true });
            await SaveJsonLogToDisk(ejecucionJson, "LogEjecucion.json");
            await SaveJsonLogToDisk(activosJson, "ActivosProcesados.json");
            await SaveJsonLogToDisk(errorJson, "LogError.json");

            await SaveJsonLogAsAttachment(_http, _siteUrl, _listTitle, itemId, "LogEjecucion.json", ejecucionPath, ct);
            await SaveJsonLogAsAttachment(_http, _siteUrl, _listTitle, itemId, "ActivosProcesados.json", activosPath, ct);
            await SaveJsonLogAsAttachment(_http, _siteUrl, _listTitle, itemId, "LogError.json", erroresPath, ct);
            }

        private async Task SaveJsonLogToDisk(string json, string logName)
            {
            var folder = Path.Combine(AppContext.BaseDirectory, "logs");

            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            var path = Path.Combine(folder, logName);
            await File.WriteAllTextAsync(path, json, Encoding.UTF8);
            }

        private async Task SaveJsonLogAsAttachment(
    HttpClient http,
    string siteUrl,
    string listTitle,
    int itemId,
    string fileName,
    string filePath,
    CancellationToken ct)
            {
            try
                {
                Console.WriteLine(filePath);
                Console.WriteLine(File.Exists(filePath));
                var bytes = await File.ReadAllBytesAsync(filePath, ct);

                var getAttachmentsUrl =
                    $"{siteUrl}/_api/web/lists/getbytitle('{listTitle}')/items({itemId})/AttachmentFiles";

                Console.WriteLine("Consultando adjuntos en SharePoint...");
                using var getResp = await http.GetAsync(getAttachmentsUrl, ct);
                if (!getResp.IsSuccessStatusCode)
                    {
                    Console.WriteLine($"❌ Error obteniendo adjuntos: {getResp.StatusCode}");
                    return;
                    }

                var content = await getResp.Content.ReadAsStringAsync(ct);
                using var doc = JsonDocument.Parse(content);

                var exists = false;

                if (doc.RootElement.TryGetProperty("value", out var arr))
                    {
                    foreach (var att in arr.EnumerateArray())
                        {
                        if (att.GetProperty("FileName").GetString()
                            .Equals(fileName, StringComparison.OrdinalIgnoreCase))
                            {
                            exists = true;
                            break;
                            }
                        }
                    }

                if (exists)
                    {
                    var deleteUrl =
                        $"{siteUrl}/_api/web/lists/getbytitle('{listTitle}')/items({itemId})/AttachmentFiles('{fileName}')";

                    var deleteReq = new HttpRequestMessage(HttpMethod.Post, deleteUrl);
                    deleteReq.Headers.Add("IF-MATCH", "*");
                    deleteReq.Headers.Add("X-HTTP-Method", "DELETE");

                    using var delResp = await http.SendAsync(deleteReq, ct);
                    if (!delResp.IsSuccessStatusCode)
                        {
                        Console.WriteLine($"❌ Error borrando adjunto {fileName}");
                        return;
                        }
                    }

                var uploadUrl =
                    $"{siteUrl}/_api/web/lists/getbytitle('{listTitle}')/items({itemId})/AttachmentFiles/add(FileName='{fileName}')";

                using var contentBytes = new ByteArrayContent(bytes);
                contentBytes.Headers.ContentType = new MediaTypeHeaderValue("application/json");

                using var uploadResp = await http.PostAsync(uploadUrl, contentBytes, ct);

                if (!uploadResp.IsSuccessStatusCode)
                    {
                    Console.WriteLine($"❌ Error subiendo {fileName}: {uploadResp.StatusCode}");
                    }
                else
                    {
                    Console.WriteLine($"✅ Log {fileName} guardado como adjunto");
                    }
                }catch (TaskCanceledException ex)
                {
                Console.WriteLine($"⚠ Timeout o cancelación al subir {fileName}");
                Console.WriteLine(ex.Message);
                }
            
            }

        }
    }
