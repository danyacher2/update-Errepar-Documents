using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Errepar.MetadataManager.Process.Services
{
    /// <summary>
    /// Helper estático para gestionar archivos adjuntos en SharePoint.
    /// Permite actualizar archivos existentes agregando contenido nuevo.
    /// </summary>
    public static class SharePointAttachmentHelper
    {
        /// <summary>
        /// Obtiene el contenido de un archivo adjunto específico por nombre.
        /// </summary>
        /// <param name="http">Cliente HTTP configurado con autenticación</param>
        /// <param name="siteUrl">URL del sitio de SharePoint</param>
        /// <param name="listTitle">Título de la lista</param>
        /// <param name="itemId">ID del item</param>
        /// <param name="fileName">Nombre del archivo a buscar</param>
        /// <param name="cancellationToken">Token de cancelación</param>
        /// <returns>Contenido del archivo en bytes, o null si no existe</returns>
        public static async Task<byte[]> GetAttachmentContentAsync(
            HttpClient http, 
            string siteUrl, 
            string listTitle, 
            int itemId, 
            string fileName, 
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("El nombre del archivo no puede estar vacío", nameof(fileName));

            // Primero, obtener la URL del archivo
            var fileUrl = await GetAttachmentUrlAsync(http, siteUrl, listTitle, itemId, fileName, cancellationToken);
            if (string.IsNullOrEmpty(fileUrl))
                return null; // Archivo no existe

            // Descargar el contenido del archivo
            var downloadEndpoint = $"{siteUrl.TrimEnd('/')}{fileUrl}";
            //var downloadEndpoint = $"{siteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('{listTitle}')/items({itemId})/AttachmentFiles/add(FileName='{Uri.EscapeDataString(fileName)}')";

            using var resp = await http.GetAsync(downloadEndpoint, cancellationToken).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode)
                return null;

            return await resp.Content.ReadAsByteArrayAsync(cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Actualiza un archivo adjunto existente agregando nuevo contenido, o lo crea si no existe.
        /// Si el archivo existe, elimina el anterior y crea uno nuevo con el contenido combinado.
        /// </summary>
        /// <param name="http">Cliente HTTP configurado con autenticación</param>
        /// <param name="siteUrl">URL del sitio de SharePoint</param>
        /// <param name="listTitle">Título de la lista</param>
        /// <param name="itemId">ID del item</param>
        /// <param name="fileName">Nombre del archivo</param>
        /// <param name="newContent">Nuevo contenido a agregar</param>
        /// <param name="appendMode">Si es true, agrega al final; si es false, reemplaza todo</param>
        /// <param name="cancellationToken">Token de cancelación</param>
        /// <returns>URL del archivo creado/actualizado</returns>
        public static async Task<string> UpdateOrCreateAttachmentAsync(
            HttpClient http,
            string siteUrl,
            string listTitle,
            int itemId,
            string fileName,
            byte[] newContent,
            bool appendMode = true,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("El nombre del archivo no puede estar vacío", nameof(fileName));
            if (newContent == null || newContent.Length == 0)
                throw new ArgumentException("El contenido no puede estar vacío", nameof(newContent));

            byte[] finalContent;

            if (appendMode)
            {
                // Intentar obtener contenido existente
                var existingContent = await GetAttachmentContentAsync(http, siteUrl, listTitle, itemId, fileName, cancellationToken);
                
                if (existingContent != null)
                {
                    Console.WriteLine($"Archivo '{fileName}' existe. Combinando contenido...");
                    
                    // Combinar contenido existente + nuevo
                    finalContent = new byte[existingContent.Length + newContent.Length];
                    Buffer.BlockCopy(existingContent, 0, finalContent, 0, existingContent.Length);
                    Buffer.BlockCopy(newContent, 0, finalContent, existingContent.Length, newContent.Length);
                    
                    // Eliminar archivo anterior
                    await DeleteAttachmentAsync(http, siteUrl, listTitle, itemId, fileName, cancellationToken);
                }
                else
                {
                    Console.WriteLine($"Archivo '{fileName}' no existe. Creando nuevo...");
                    finalContent = newContent;
                }
            }
            else
            {
                // Modo reemplazo: eliminar si existe y crear con nuevo contenido
                var exists = await AttachmentExistsAsync(http, siteUrl, listTitle, itemId, fileName, cancellationToken);
                if (exists)
                {
                    Console.WriteLine($"Archivo '{fileName}' existe. Reemplazando...");
                    await DeleteAttachmentAsync(http, siteUrl, listTitle, itemId, fileName, cancellationToken);
                }
                else
                {
                    Console.WriteLine($"Archivo '{fileName}' no existe. Creando nuevo...");
                }
                finalContent = newContent;
            }

            // Crear/subir el archivo con el contenido final
            return await AddAttachmentAsync(http, siteUrl, listTitle, itemId, fileName, finalContent, cancellationToken);
        }

        /// <summary>
        /// Verifica si existe un archivo adjunto con el nombre especificado.
        /// </summary>
        public static async Task<bool> AttachmentExistsAsync(
            HttpClient http,
            string siteUrl,
            string listTitle,
            int itemId,
            string fileName,
            CancellationToken cancellationToken = default)
        {
            var url = await GetAttachmentUrlAsync(http, siteUrl, listTitle, itemId, fileName, cancellationToken);
            return !string.IsNullOrEmpty(url);
        }

        // ===== Métodos privados auxiliares =====

        public static async Task<string> GetAttachmentUrlAsync(
            HttpClient http,
            string siteUrl,
            string listTitle,
            int itemId,
            string fileName,
            CancellationToken cancellationToken)
        {
            var endpoint = $"{siteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('{listTitle}')/items({itemId})/AttachmentFiles";

            using var resp = await http.GetAsync(endpoint, cancellationToken).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode)
                return null;

            var json = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            using var doc = JsonDocument.Parse(json);

            JsonElement arrayElement;
            if (doc.RootElement.TryGetProperty("d", out var d) && d.TryGetProperty("results", out arrayElement))
            {
                // ok
            }
            else if (doc.RootElement.TryGetProperty("value", out arrayElement))
            {
                // ok
            }
            else
            {
                return null;
            }

            // Buscar el archivo por nombre
            foreach (var attachment in arrayElement.EnumerateArray())
            {
                if (attachment.TryGetProperty("FileName", out var fileNameProp) &&
                    fileNameProp.GetString()?.Equals(fileName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    if (attachment.TryGetProperty("ServerRelativeUrl", out var urlProp))
                    {
                        return urlProp.GetString();
                    }
                }
            }

            return null;
        }

        private static async Task<string> AddAttachmentAsync(
            HttpClient http,
            string siteUrl,
            string listTitle,
            int itemId,
            string fileName,
            byte[] fileContent,
            CancellationToken cancellationToken)
        {
            var endpoint = $"{siteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('{listTitle}')/items({itemId})/AttachmentFiles/add(FileName='{Uri.EscapeDataString(fileName)}')";

            using var request = new HttpRequestMessage(HttpMethod.Post, endpoint);

            // Configurar contenido
            request.Content = new ByteArrayContent(fileContent);

            // IMPORTANTE: Verificar que el HttpClient tenga el token configurado
            if (http.DefaultRequestHeaders.Authorization == null)
            {
                Console.WriteLine("⚠️ ADVERTENCIA: HttpClient no tiene Authorization header configurado");
                throw new InvalidOperationException("El HttpClient debe tener configurado el Authorization header (Bearer token)");
            }
            else
            {
                Console.WriteLine($"✓ Authorization header presente: {http.DefaultRequestHeaders.Authorization.Scheme} {http.DefaultRequestHeaders.Authorization.Parameter?.Substring(0, Math.Min(20, http.DefaultRequestHeaders.Authorization.Parameter?.Length ?? 0))}...");
            }

            // Headers requeridos por SharePoint REST API
            request.Headers.Accept.Clear();
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.Add("Accept", "application/json;odata=verbose");

            // Content-Type para archivos binarios
            request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            request.Content.Headers.ContentLength = fileContent.Length;

            Console.WriteLine($"Subiendo archivo a: {endpoint}");
            using var resp = await http.SendAsync(request, cancellationToken).ConfigureAwait(false);

            if (!resp.IsSuccessStatusCode)
            {
                var errorContent = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

                // Mostrar headers de autenticación esperados si es 401
                if (resp.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    Console.WriteLine("\n⚠️ ERROR 401 - NO AUTORIZADO");
                    Console.WriteLine("El token puede estar expirado o no tener los permisos necesarios");
                    if (resp.Headers.WwwAuthenticate != null)
                    {
                        Console.WriteLine("WWW-Authenticate headers:");
                        foreach (var auth in resp.Headers.WwwAuthenticate)
                        {
                            Console.WriteLine($"  {auth}");
                        }
                    }
                }

                throw new HttpRequestException(
                    $"Error al agregar adjunto.\n" +
                    $"Status: {resp.StatusCode}\n" +
                    $"Endpoint: {endpoint}\n" +
                    $"FileName: {fileName}\n" +
                    $"FileSize: {fileContent.Length} bytes\n" +
                    $"Response: {errorContent}");
            }

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

        private static async Task DeleteAttachmentAsync(
            HttpClient http,
            string siteUrl,
            string listTitle,
            int itemId,
            string fileName,
            CancellationToken cancellationToken)
        {
            var endpoint = $"{siteUrl.TrimEnd('/')}/_api/web/lists/getbytitle('{listTitle}')/items({itemId})/AttachmentFiles/getByFileName('{Uri.EscapeDataString(fileName)}')";

            var request = new HttpRequestMessage(HttpMethod.Delete, endpoint);
            request.Headers.Add("IF-MATCH", "*");

            using var resp = await http.SendAsync(request, cancellationToken).ConfigureAwait(false);
            
            // 200 OK o 204 No Content son éxito
            if (!resp.IsSuccessStatusCode && resp.StatusCode != System.Net.HttpStatusCode.NoContent)
            {
                var errorContent = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                throw new HttpRequestException($"Error al eliminar adjunto: {errorContent}");
            }
        }
    }
}
