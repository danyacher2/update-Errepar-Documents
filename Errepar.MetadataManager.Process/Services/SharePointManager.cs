using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Errepar.MetadataManager.Process.Models;

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
                           "?$select=Id,Title,Cambios,Cant_x0020_Activos_x0020_Procesados,Cant_x0020_Activos_x0020_Seleccionados,Created,Author/Title,Editor/Title,EstadoProceso,FechaFinalizado,FechaPendiente,Link,Modified" +
                           "&$expand=Author,Editor" +
                           "&$filter=(EstadoProceso eq 'Pendiente') or (EstadoProceso eq 'En Pausa')";

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
                it.Titulo = TryGetString(el, "Title", "Titulo");
                it.Cambios = TryGetString(el, "Cambios");
                it.CantActivosProcesados = TryGetInt(el, "Cant_x0020_Activos_x0020_Procesados", "CantActivosProcesados");
                it.CantActivosSeleccionados = TryGetInt(el, "Cant_x0020_Activos_x0020_Seleccionados", "CantActivosSeleccionados");
                it.Creado = TryGetDateTime(el, "Created", "Creado");
                it.Modificado = TryGetDateTime(el, "Modified", "Modificado");
                it.FechaFinalizado = TryGetDateTime(el, "FechaFinalizado");
                it.FechaPendiente = TryGetDateTime(el, "FechaPendiente");
                it.Link = TryGetLink(el, "Link");
                it.EstadoProceso = TryGetString(el, "EstadoProceso", "Estado_x0020_Proceso");
                it.EjecutadoPor = TryGetNestedUser(el, "ExecutedBy", "EjecutadoPor"); // fallback
                it.CreadoPor = TryGetNestedUser(el, "Author", "CreatedBy");
                it.ModificadoPor = TryGetNestedUser(el, "Editor", "ModifiedBy");

                items.Add(it);
            }

            return items;
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
    }
}