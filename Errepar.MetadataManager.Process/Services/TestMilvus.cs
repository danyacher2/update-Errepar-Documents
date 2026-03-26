using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Errepar.MetadataManager.Process.Services
{
    public class TestMilvus
    {
        private readonly HttpClient _http;
        private readonly string _authEndpoint;
        private readonly string _milvusEndpoint;
        private readonly string _apiKey;
        private string? _cachedToken;
        private DateTime _tokenExpiration;

        public TestMilvus(HttpClient? httpClient = null)
        {
            _http = httpClient ?? new HttpClient();
            _authEndpoint = "https://accounts.uat.errepar.com/syserrepar/integration/authenticationandauthorization/e-auth-api/auth/job/getJobCredentialsEAuth";
            _milvusEndpoint = "https://api.uat.errepar.com/syserrepar/embeddeddocumentai/searchByTimestamp";
            _apiKey = "e5169fe4-1c39-4356-b148-1f210cc2431b";
        }

        /// <summary>
        /// Obtiene metadatos de Milvus filtrando por timestamp del documento
        /// </summary>
        /// <param name="timestamp">Timestamp del documento (formato: yyyyMMddHHmmssfff, ej: 20240326054746556)</param>
        /// <param name="cancellationToken">Token de cancelación</param>
        /// <returns>Respuesta JSON del servicio Milvus</returns>
        public async Task<JsonDocument?> GetMetadatosMilvusAsync(string timestamp, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(timestamp))
                throw new ArgumentException("El timestamp no puede estar vacío", nameof(timestamp));

            try
            {
                // 1. Obtener token de autenticación
                var token = await GetAuthTokenAsync(cancellationToken).ConfigureAwait(false);

                // 2. Construir URL con query params
                var requestUrl = $"{_milvusEndpoint}?timestamp={timestamp}";

                // 3. Configurar request con el token
                using var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // 4. Ejecutar petición
                Console.WriteLine($"Consultando Milvus con timestamp: {timestamp}");
                using var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
                
                // 5. Validar respuesta
                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                    throw new HttpRequestException(
                        $"Error al consultar Milvus. Status: {response.StatusCode}, Body: {errorContent}");
                }

                // 6. Parsear y retornar JSON
                var jsonContent = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                var jsonDoc = JsonDocument.Parse(jsonContent);

                Console.WriteLine($"Datos obtenidos de Milvus: {jsonContent.Length} caracteres");
                return jsonDoc;
            }
            catch (HttpRequestException ex)
            {
                Console.Error.WriteLine($"Error HTTP al consultar Milvus: {ex.Message}");
                throw;
            }
            catch (JsonException ex)
            {
                Console.Error.WriteLine($"Error al parsear JSON de Milvus: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error inesperado al consultar Milvus: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Obtiene el token de autenticación desde el servicio e-auth-api.
        /// Implementa cache simple para evitar múltiples llamadas.
        /// CORREGIDO: Ahora sigue el mismo patrón que MilvusManager.cs
        /// </summary>
        private async Task<string> GetAuthTokenAsync(CancellationToken cancellationToken)
        {
            // Retornar token en cache si aún es válido (con 5 min de margen)
            if (!string.IsNullOrEmpty(_cachedToken) && DateTime.UtcNow < _tokenExpiration.AddMinutes(-5))
            {
                Console.WriteLine("Usando token en cache");
                return _cachedToken;
            }

            try
            {
                Console.WriteLine("Obteniendo nuevo token de autenticación...");

                // Construir body JSON como string (igual que en MilvusManager.cs)
                string bodyTokenMilvus = "{\"key\":\"" + _apiKey + "\"}";

                // Crear HttpRequestMessage con headers explícitos (igual que MilvusManager.cs)
                var httpRequestToken = new HttpRequestMessage
                {
                    Method = HttpMethod.Post,
                    RequestUri = new Uri(_authEndpoint),
                    Headers = {
                        { HttpRequestHeader.ContentType.ToString(), "application/json" },
                        { HttpRequestHeader.Accept.ToString(), "application/json" }
                    },
                    Content = new StringContent(bodyTokenMilvus, Encoding.UTF8, "application/json")
                };

                // Ejecutar request
                var response = await _http.SendAsync(httpRequestToken, cancellationToken).ConfigureAwait(false);

                // Validar status Found (302) - ESTE ES EL COMPORTAMIENTO CORRECTO
                if (response.StatusCode != HttpStatusCode.Found)
                {
                    var errorContent = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                    throw new HttpRequestException(
                        $"Error al obtener token. Status esperado: Found (302), recibido: {response.StatusCode}. Body: {errorContent}");
                }

                // Parsear respuesta y extraer jwt2 (NO "token" ni "access_token")
                var responseBody = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                
                using var jsonDoc = JsonDocument.Parse(responseBody);
                
                // Extraer jwt2
                if (!jsonDoc.RootElement.TryGetProperty("jwt2", out var jwt2Prop))
                {
                    throw new InvalidOperationException(
                        $"No se encontró la propiedad 'jwt2' en la respuesta. JSON: {responseBody}");
                }

                var jwt2 = jwt2Prop.GetString();
                
                if (string.IsNullOrEmpty(jwt2))
                {
                    throw new InvalidOperationException("El token jwt2 está vacío en la respuesta");
                }

                // Decodificar el JWT para obtener la expiración (opcional pero recomendado)
                // Por ahora, asumimos 12 horas de validez (43200 segundos)
                int expiresIn = 43200;

                // Cachear token y establecer expiración
                _cachedToken = jwt2;
                _tokenExpiration = DateTime.UtcNow.AddSeconds(expiresIn);

                Console.WriteLine($"Token jwt2 obtenido correctamente. Expira en {expiresIn} segundos");
                return jwt2;
            }
            catch (HttpRequestException ex)
            {
                Console.Error.WriteLine($"Error HTTP al obtener token: {ex.Message}");
                throw;
            }
            catch (JsonException ex)
            {
                Console.Error.WriteLine($"Error al parsear JSON del token: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error inesperado al obtener token: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Sobrecarga para obtener metadatos con timestamp como long
        /// </summary>
        public async Task<JsonDocument?> GetMetadatosMilvusAsync(long timestamp, CancellationToken cancellationToken = default)
        {
            return await GetMetadatosMilvusAsync(timestamp.ToString(), cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Sobrecarga para obtener metadatos con timestamp como DateTime
        /// </summary>
        public async Task<JsonDocument?> GetMetadatosMilvusAsync(DateTime timestamp, CancellationToken cancellationToken = default)
        {
            // Convertir DateTime a formato yyyyMMddHHmmssfff
            var timestampStr = timestamp.ToString("yyyyMMddHHmmssfff");
            return await GetMetadatosMilvusAsync(timestampStr, cancellationToken).ConfigureAwait(false);
        }
    }
}