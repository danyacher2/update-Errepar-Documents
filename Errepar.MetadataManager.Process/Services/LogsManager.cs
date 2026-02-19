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

    public class ItemLogActivosProcesados
    {
        public int ItemId { get; set; }
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

        public async Task SaveLogEjecucion(ItemLogEjecucion item)
        {
            _logEjecucionMemoria.Add(item);

            var json = JsonSerializer.Serialize(item) + Environment.NewLine;
            await File.AppendAllTextAsync("logs/LogEjecucion.ndjson", json);
        }


        public async Task SaveLogActivosProcesados(ItemLogActivosProcesados item)
        {
            _logActivosMemoria.Add(item);

            var json = JsonSerializer.Serialize(item) + Environment.NewLine;
            await File.AppendAllTextAsync("logs/ActivosProcesados.ndjson", json);
        }



        public async Task SyncLogs( int itemId, CancellationToken ct)
        {
            // Logs en memoria
            var ejecucionJson = JsonSerializer.Serialize(_logEjecucionMemoria, new JsonSerializerOptions { WriteIndented = true });
            var activosJson = JsonSerializer.Serialize(_logActivosMemoria, new JsonSerializerOptions { WriteIndented = true });

            // Guardar en disco
            await SaveJsonLogToDisk(ejecucionJson, "LogEjecucion.json");
            await SaveJsonLogToDisk(activosJson, "ActivosProcesados.json");

            // Subir como adjuntos
            await SaveJsonLogAsAttachment(_http, _siteUrl, _listTitle, itemId, "LogEjecucion.json", ejecucionJson, ct);
            await SaveJsonLogAsAttachment(_http, _siteUrl, _listTitle, itemId, "ActivosProcesados.json", activosJson, ct);
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
            string json,
            CancellationToken ct)
        {
            var bytes = Encoding.UTF8.GetBytes(json);

            // 1) Obtener adjuntos actuales
            var getAttachmentsUrl =
                $"{siteUrl}/_api/web/lists/getbytitle('{listTitle}')/items({itemId})/AttachmentFiles";

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

            // 2) Si existe → borrar
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

            // 3) Subir adjunto
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
        }
    }
}
