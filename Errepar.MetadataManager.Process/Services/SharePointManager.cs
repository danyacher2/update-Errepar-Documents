using Errepar.MetadataManager.Process.Models;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Errepar.MetadataManager.Process.Services
{
    public class SharePointManager
    {
        private readonly HttpClient _http;
        private readonly string _siteUrl;
        private readonly string _listTitle;

        // Construir con HttpClient (puede inyectarse) y la url de sitio SharePoint.
        // authToken: bearer token si usas OAuth. Si usas otro método, ajusta la cabecera/autenticación.
        public SharePointManager(HttpClient httpClient, string siteUrl, string listTitle = "MetadataManager", string authToken = null)
        {
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _siteUrl = siteUrl?.TrimEnd('/') ?? throw new ArgumentNullException(nameof(siteUrl));
            _listTitle = listTitle;

            if (!string.IsNullOrEmpty(authToken))
                _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);
            _http.DefaultRequestHeaders.Accept.Clear();
            _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        // Recupera todos los items cuyo EstadoProceso sea "Pendiente" o "En Pausa".
        // Nota: Ajusta los nombres internos de los campos ($select) si en tu lista difieren.
        public async Task<List<ItemMetadataManager>> GetPendingOrPausedAsync(CancellationToken cancellationToken = default)
        {
            var items = new List<ItemMetadataManager>();

            // Endpoint REST; se usa odata v4 ($select + $expand).
            // Ajustar internal names si los tuyos difieren (p. ej. Estado_x0020_Proceso).
            var endpoint = $"{_siteUrl}/_api/web/lists/getbytitle('{_listTitle}')/items" +
                "?$select=Id,Title,Cambios,CantActivosProcesados,Scope,CantActivosSeleccionados,Created," +
                "Author/Title,Editor/Title,EstadoProceso,FechaFinalizado,FechaPendiente,LinkMetadataManager," +
                "Modified,EjecutadoPor/Title,EjecutadoPor/Id,AttachmentFiles/ServerRelativeUrl" +
                "&$expand=Author,Editor,AttachmentFiles,EjecutadoPor" +
                "&$filter=(EstadoProceso eq 'Pendiente')";
                //"&$filter=(EstadoProceso eq 'Pendiente') or (EstadoProceso eq 'En Pausa')";
            ;
             using var resp = await _http.GetAsync(endpoint, cancellationToken).ConfigureAwait(false);
             resp.EnsureSuccessStatusCode();

             var json = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

            using var doc = JsonDocument.Parse(json);

            // Manejo flexible: SharePoint puede devolver { "value": [...] } o { "d": { "results": [...] } }
            JsonElement arrayElement;
            if (doc.RootElement.TryGetProperty("value", out arrayElement))
            {
                // ok
            }
            else if (doc.RootElement.TryGetProperty("d", out var dEl) && dEl.TryGetProperty("results", out arrayElement))
            {
                // ok
            }
            else
            {
                // fallback: intentar root como array
                if (doc.RootElement.ValueKind == JsonValueKind.Array)
                {
                    arrayElement = doc.RootElement;
                }
                else
                {
                    return items; // no encontrado
                }
            }

            foreach (var el in arrayElement.EnumerateArray())
            {
                var it = new ItemMetadataManager();

                it.Id = TryGetInt(el, "Id") ?? 0;
                it.Cambios = TryGetJson(el, "Cambios");
                

                it.CantActivosProcesados = TryGetInt(el, "CantActivosProcesados", "Cant Activos Procesados");
                it.CantActivosSeleccionados = TryGetInt(el, "CantActivosSeleccionados", "Cant Activos Seleccionados");
                it.Creado = TryGetDateTime(el, "Created", "Creado");
                it.EjecutadoPor = TryGetNestedUser(el, "EjecutadoPor", "Ejecutado Por"); // fallback
                it.EstadoProceso = TryGetString(el, "EstadoProceso", "Estado Proceso");
                it.FechaFinalizado = TryGetDateTime(el, "FechaFinalizado", "Fecha Finalizado");
                it.FechaPendiente = TryGetDateTime(el, "FechaPendiente", "Fecha Pendiente");
                it.Link = TryGetLink(el, "LinkMetadataManager");
                it.Modificado = TryGetDateTime(el, "Modified", "Modificado");
                it.Titulo = TryGetString(el, "Title", "Titulo");
                it.CreadoPor = TryGetNestedUser(el, "Author", "CreatedBy");
                it.ModificadoPor = TryGetNestedUser(el, "Editor", "ModifiedBy");
                it.Scope = TryGetNestedUser(el, "Scope", "Scope");
                it.Adjuntos = TryGetAttachments(el, "AttachmentFiles");
                it.Activos = (await TryGetActivosAsync(_http, _siteUrl, _listTitle, it.Id, it.Adjuntos, cancellationToken)).ToList();

                items.Add(it);
            }

            return items;
        }

        public async Task UpdateListItemAsync(int id, Dictionary<string, object> fields)
        {
            var url = $"{_siteUrl}/_api/web/lists/getbytitle('{_listTitle}')/items({id})";

            var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Add("IF-MATCH", "*");
            request.Headers.Add("X-HTTP-Method", "MERGE");

            var json = JsonSerializer.Serialize(fields);

            request.Content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _http.SendAsync(request);

            response.EnsureSuccessStatusCode();
        }
        private static JsonDocument TryGetJson(JsonElement el, string fieldName)
    {
        if (!el.TryGetProperty(fieldName, out var prop))
            return null;

        if (prop.ValueKind == JsonValueKind.Null)
            return null;

        // 🟢 Caso 1: ya es JSON real
        if (prop.ValueKind == JsonValueKind.Array || prop.ValueKind == JsonValueKind.Object)
        {
            return JsonDocument.Parse(prop.GetRawText());
        }

        // 🟡 Caso 2: string
        if (prop.ValueKind == JsonValueKind.String)
        {
            var raw = prop.GetString();

            if (string.IsNullOrWhiteSpace(raw))
                return null;

            raw = raw.Trim();

            // 🔴 Si viene HTML → limpiamos tags
            if (raw.StartsWith("<"))
            {
                // 1) quitar tags HTML
                raw = Regex.Replace(raw, "<.*?>", string.Empty);

                // 2) decodificar entidades HTML
                raw = WebUtility.HtmlDecode(raw);

                raw = raw.Trim();
            }

            // 🧠 ahora intentamos parsear como JSON real
            if ((raw.StartsWith("[") && raw.EndsWith("]")) ||
                (raw.StartsWith("{") && raw.EndsWith("}")))
            {
                try
                {
                    return JsonDocument.Parse(raw);
                }
                catch (JsonException)
                {
                    // si está corrupto, lo envolvemos como string
                    var safe = JsonSerializer.Serialize(raw);
                    return JsonDocument.Parse(safe);
                }
            }

            // 🔵 No es JSON → lo devolvemos como string JSON
            var safeJson = JsonSerializer.Serialize(raw);
            return JsonDocument.Parse(safeJson);
        }

        return null;
    }
    /// <summary>
    /// Busca y descarga el contenido del archivo "Activos" de los adjuntos.
    /// Parsea el contenido y lo devuelve como array de strings.
    /// </summary>
    private async Task<string[]> TryGetActivosAsync(HttpClient http, string siteUrl, string listTitle, int itemId, string datosAdjuntos, CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(datosAdjuntos))
                return Array.Empty<string>();

            // Buscar archivo que contenga "Activos" en su nombre (case-insensitive)
            var urls = datosAdjuntos.Split(';', StringSplitOptions.RemoveEmptyEntries);
            string archivoActivosUrl = null;
            string fileName = null;

            foreach (var url in urls)
            {
                fileName = System.IO.Path.GetFileName(url);
                if (fileName.Contains("Activos", StringComparison.OrdinalIgnoreCase))
                {
                    archivoActivosUrl = url;
                    break;
                }
            }

            if (string.IsNullOrEmpty(archivoActivosUrl))
            {
                Console.WriteLine($"⚠️ No se encontró archivo 'Activos' en item {itemId}");
                return Array.Empty<string>();
            }

            try
            {
                var serverRelativeUrl = archivoActivosUrl; // ya lo tenés
                var downloadEndpoint = $"{siteUrl}/_api/web/GetFileByServerRelativeUrl('{serverRelativeUrl}')/$value";

                using var resp = await http.GetAsync(downloadEndpoint, HttpCompletionOption.ResponseHeadersRead, cancellationToken);

                if (!resp.IsSuccessStatusCode)
                {
                    Console.WriteLine($"❌ Error al descargar archivo: Status {resp.StatusCode}");
                    return Array.Empty<string>();
                }

                var contenido = await resp.Content.ReadAsByteArrayAsync(cancellationToken).ConfigureAwait(false);               

                // Convertir bytes a texto
                var texto = System.Text.Encoding.UTF8.GetString(contenido);
                using var doc = JsonDocument.Parse(texto);

                var root = doc.RootElement;

                if (!root.TryGetProperty("documents", out var documentsElement))
                {
                    Console.WriteLine("❌ El JSON no tiene la propiedad 'Documents'");
                    return Array.Empty<string>();
                }

                if (documentsElement.ValueKind != JsonValueKind.Array)
                {
                    Console.WriteLine("❌ 'Documents' no es un array");
                    return Array.Empty<string>();
                }

                var activos = new List<string>();

                foreach (var item in documentsElement.EnumerateArray())
                {
                    // ejemplo: { "id": "...", "nombre": "..." }
                    if (item.TryGetProperty("id", out var idProp))
                        activos.Add(idProp.GetString());
                    else if (item.TryGetProperty("nombre", out var nombreProp))
                        activos.Add(nombreProp.GetString());
                }

                return activos.ToArray();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error al obtener activos del item {itemId}: {ex.Message}");
                Console.WriteLine(ex);
                return Array.Empty<string>();
            }
        }

        /// <summary>
        /// Parsea contenido de texto plano (una línea por activo)
        /// </summary>
        private static Array ParseTextActivos(string texto)
        {
            var lineas = texto.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var activos = lineas.Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
            Console.WriteLine($"✅ {activos.Length} activos encontrados (texto)");
            return activos;
        }

        /// <summary>
        /// Parsea contenido JSON (espera un array de strings o array de objetos con propiedad "nombre" o "id")
        /// </summary>
        private static Array ParseJsonActivos(string json)
        {
            try
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                // Si es un array directo
                if (root.ValueKind == JsonValueKind.Array)
                {
                    var activos = new List<string>();
                    foreach (var elemento in root.EnumerateArray())
                    {
                        if (elemento.ValueKind == JsonValueKind.String)
                        {
                            // Array de strings: ["activo1", "activo2"]
                            activos.Add(elemento.GetString());
                        }
                        else if (elemento.ValueKind == JsonValueKind.Object)
                        {
                            // Array de objetos: [{"id": "123", "nombre": "activo1"}]
                            var valor = TryGetString(elemento, "id", "nombre", "codigo", "activo");
                            if (!string.IsNullOrEmpty(valor))
                                activos.Add(valor);
                        }
                    }
                    Console.WriteLine($"✅ {activos.Count} activos encontrados (JSON array)");
                    return activos.ToArray();
                }
                // Si tiene una propiedad "activos" con el array
                else if (root.TryGetProperty("activos", out var activosArray) && activosArray.ValueKind == JsonValueKind.Array)
                {
                    var activos = new List<string>();
                    foreach (var elemento in activosArray.EnumerateArray())
                    {
                        if (elemento.ValueKind == JsonValueKind.String)
                            activos.Add(elemento.GetString());
                        else if (elemento.ValueKind == JsonValueKind.Object)
                        {
                            var valor = TryGetString(elemento, "id", "nombre", "codigo", "activo");
                            if (!string.IsNullOrEmpty(valor))
                                activos.Add(valor);
                        }
                    }
                    Console.WriteLine($"✅ {activos.Count} activos encontrados (JSON objeto)");
                    return activos.ToArray();
                }

                Console.WriteLine("⚠️ Formato JSON no reconocido, retornando array vacío");
                return Array.Empty<string>();
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"❌ Error al parsear JSON: {ex.Message}");
                return Array.Empty<string>();
            }
        }

        // ----- Helpers de parseo JSON -----
        private static string TryGetString(JsonElement el, params string[] names)
        {
            foreach (var name in names)
            {
                if (el.TryGetProperty(name, out var p) && p.ValueKind != JsonValueKind.Null)
                    return p.GetString();
            }
            return null;
        }

        private static int? TryGetInt(JsonElement el, params string[] names)
        {
            foreach (var name in names)
            {
                if (el.TryGetProperty(name, out var p) && (p.ValueKind == JsonValueKind.Number))
                {
                    if (p.TryGetInt32(out var v)) return v;
                }
                else if (p.ValueKind == JsonValueKind.String)
                {
                    if (int.TryParse(p.GetString(), out var v)) return v;
                }
            }
            return null;
        }

        private static DateTime? TryGetDateTime(JsonElement el, params string[] names)
        {
            foreach (var name in names)
            {
                if (el.TryGetProperty(name, out var p) && p.ValueKind != JsonValueKind.Null)
                {
                    if (p.ValueKind == JsonValueKind.String && DateTime.TryParse(p.GetString(), out var dt))
                        return dt;
                    if (p.ValueKind == JsonValueKind.Number && p.TryGetInt64(out var epoch))
                        return DateTimeOffset.FromUnixTimeSeconds(epoch).DateTime;
                }
            }
            return null;
        }

        private static string TryGetLink(JsonElement el, string name)
        {
            if (el.TryGetProperty(name, out var p) && p.ValueKind == JsonValueKind.Object)
            {
                if (p.TryGetProperty("Url", out var url) && url.ValueKind == JsonValueKind.String)
                    return url.GetString();
                if (p.TryGetProperty("url", out var url2) && url2.ValueKind == JsonValueKind.String)
                    return url2.GetString();
            }
            else if (el.TryGetProperty(name, out p) && p.ValueKind == JsonValueKind.String)
            {
                return p.GetString();
            }
            return null;
        }

        private static string TryGetNestedUser(JsonElement el, params string[] names)
        {
            foreach (var name in names)
            {
                if (el.TryGetProperty(name, out var p) && p.ValueKind == JsonValueKind.Object)
                {
                    if (p.TryGetProperty("Title", out var t) && t.ValueKind == JsonValueKind.String)
                        return t.GetString();
                    if (p.TryGetProperty("LoginName", out var ln) && ln.ValueKind == JsonValueKind.String)
                        return ln.GetString();
                }
            }
            return null;
        }

        private static string TryGetAttachments(JsonElement el, string name)
        {
            if (!el.TryGetProperty(name, out var attachmentsProperty))
                return null;

            // AttachmentFiles puede venir como objeto con "results" o directamente como array
            JsonElement attachmentsArray;
            if (attachmentsProperty.ValueKind == JsonValueKind.Object &&
                attachmentsProperty.TryGetProperty("results", out attachmentsArray))
            {
                // ok
            }
            else if (attachmentsProperty.ValueKind == JsonValueKind.Array)
            {
                attachmentsArray = attachmentsProperty;
            }
            else
            {
                return null;
            }

            var urls = new List<string>();
            foreach (var attachment in attachmentsArray.EnumerateArray())
            {
                // Los adjuntos tienen propiedades como ServerRelativeUrl o FileName
                if (attachment.TryGetProperty("ServerRelativeUrl", out var urlProp) &&
                    urlProp.ValueKind == JsonValueKind.String)
                {
                    urls.Add(urlProp.GetString());
                }
            }

            return urls.Count > 0 ? string.Join(";", urls) : null;
        }

        public async Task<string> AddAttachmentAsync(int itemId, string fileName, byte[] fileContent, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("El nombre del archivo no puede estar vacío", nameof(fileName));
            if (fileContent == null || fileContent.Length == 0)
                throw new ArgumentException("El contenido del archivo no puede estar vacío", nameof(fileContent));

            // Endpoint para agregar adjunto
            var endpoint = $"{_siteUrl}/_api/web/lists/getbytitle('{_listTitle}')/items({itemId})/AttachmentFiles/add(FileName='{Uri.EscapeDataString(fileName)}')";

            using var content = new ByteArrayContent(fileContent);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

            using var resp = await _http.PostAsync(endpoint, content, cancellationToken).ConfigureAwait(false);

            if (!resp.IsSuccessStatusCode)
            {
                var errorContent = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                throw new HttpRequestException(
                    $"Error al agregar adjunto a SharePoint.\n" +
                    $"Status: {resp.StatusCode} ({(int)resp.StatusCode})\n" +
                    $"ItemId: {itemId}, FileName: {fileName}\n" +
                    $"Response: {errorContent}");
            }

            resp.EnsureSuccessStatusCode();

            var json = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            using var doc = JsonDocument.Parse(json);

            if (doc.RootElement.TryGetProperty("d", out var d) && 
                d.TryGetProperty("ServerRelativeUrl", out var url))
            {
                return url.GetString();
            }
            else if (doc.RootElement.TryGetProperty("ServerRelativeUrl", out url))
            {
                return url.GetString();
            }

            return null;
        }
        public async Task<List<string>> AddAttachmentsAsync(int itemId, Dictionary<string, byte[]> attachments, CancellationToken cancellationToken = default)
        {
            if (attachments == null || attachments.Count == 0)
                throw new ArgumentException("Debe proporcionar al menos un adjunto", nameof(attachments));

            var uploadedUrls = new List<string>();

            foreach (var attachment in attachments)
            {
                try
                {
                    var url = await AddAttachmentAsync(itemId, attachment.Key, attachment.Value, cancellationToken).ConfigureAwait(false);
                    if (!string.IsNullOrEmpty(url))
                    {
                        uploadedUrls.Add(url);
                        Console.WriteLine($"✓ Adjunto agregado: {attachment.Key} -> {url}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"✗ Error al agregar adjunto '{attachment.Key}': {ex.Message}");
                    // Continuar con los siguientes archivos
                }
            }

            return uploadedUrls;
        }

     
        public async Task<List<string>> GetAttachmentsAsync(int itemId, CancellationToken cancellationToken = default)
        {
            var endpoint = $"{_siteUrl}/_api/web/lists/getbytitle('{_listTitle}')/items({itemId})/AttachmentFiles";

            using var resp = await _http.GetAsync(endpoint, cancellationToken).ConfigureAwait(false);
            resp.EnsureSuccessStatusCode();

            var json = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            using var doc = JsonDocument.Parse(json);

            var urls = new List<string>();
            JsonElement arrayElement;

            if (doc.RootElement.TryGetProperty("d", out var d) && 
                d.TryGetProperty("results", out arrayElement))
            {
                // ok
            }
            else if (doc.RootElement.TryGetProperty("value", out arrayElement))
            {
                // ok
            }
            else
            {
                return urls;
            }

            foreach (var attachment in arrayElement.EnumerateArray())
            {
                if (attachment.TryGetProperty("ServerRelativeUrl", out var urlProp) && 
                    urlProp.ValueKind == JsonValueKind.String)
                {
                    urls.Add(urlProp.GetString());
                }
            }

            return urls;
        }
    }
}