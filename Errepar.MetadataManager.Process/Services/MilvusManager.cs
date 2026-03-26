using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Errepar.MetadataManager.Process.Services
{
    // Cliente mínimo para Milvus REST/SDK. Normalmente Milvus se consume con su SDK (gRPC/protobuf).
    // Aquí se deja un esqueleto con llamadas HTTP para ilustrar integración vía REST proxy si existiera.
    public class MilvusManager
    {
        
        private readonly HttpClient _http;
        private readonly string _milvusBaseUrl;

        public static async Task<(bool ok, string msg)> EnviarAMilvus(HttpClient client, string json, int id, bool isIA = false)
        {
            //string bodyTokenMilvus = "{\"key\":\"e5169fe4-1c39-4356-b148-1f210cc2431b\"}";
            //string urlTokenMilvus = "https://accounts.errepar.com/eauth/auth/job/getJobCredentialsEAuth";
            const int maxIntentos = 3;
            const int delayEntreIntentos = 5000; // 5 segundos

            int intentoActual = 0;
            (bool ok, string msg) resultado = (false, string.Empty);

            while (intentoActual < maxIntentos)
            {
                intentoActual++;

                try
                {
                    Console.WriteLine($"🔄 MilvusManager - Intento {intentoActual}/{maxIntentos} para elemento {id}");

                    resultado = await EnviarAMilvusInterno(client, json, id, isIA);

                    if (resultado.ok)
                    {
                        if (intentoActual > 1)
                        {
                            Console.WriteLine($"✔️ Milvus exitoso en intento {intentoActual}/{maxIntentos} para elemento {id}");
                        }
                        return resultado;
                    }

                    Console.WriteLine($"⚠️ MilvusManager - Error en intento {intentoActual}/{maxIntentos}: {resultado.msg}");

                    if (intentoActual < maxIntentos)
                    {
                        await Task.Delay(delayEntreIntentos);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ MilvusManager - Excepción en intento {intentoActual}/{maxIntentos}:");
                    Console.WriteLine($"   Mensaje: {ex.Message}");
                    Console.WriteLine($"   Tipo: {ex.GetType().Name}");
                    if (ex.InnerException != null)
                    {
                        Console.WriteLine($"   Inner Exception: {ex.InnerException.Message}");
                    }

                    if (intentoActual >= maxIntentos)
                    {
                        return (false, $"Excepción después de {maxIntentos} intentos: {ex.Message}");
                    }

                    await Task.Delay(delayEntreIntentos);
                }
            }

            return (false, $"Falló después de {maxIntentos} intentos. Último error: {resultado.msg}");
        }

        private static async Task<(bool ok, string msg)> EnviarAMilvusInterno(HttpClient client, string json, int id, bool isIA)
        {
            //string bodyTokenMilvus = "{\"key\":\"e5169fe4-1c39-4356-b148-1f210cc2431b\"}";
            string bodyTokenMilvus = "{\"key\":\"e5169fe4-1c39-4356-b148-1f210cc2431b\"}";
            string urlTokenMilvus = "https://accounts.uat.errepar.com/syserrepar/integration/authenticationandauthorization/e-auth-api/auth/job/getJobCredentialsEAuth";
            //string urlTokenMilvus = "https://accounts.uat.errepar.com/eauth/auth/job/getJobCredentialsEAuth";

            var httpRequestToken = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(urlTokenMilvus),
                Headers = {
                  { HttpRequestHeader.ContentType.ToString(), "application/json" },
                  { HttpRequestHeader.Accept.ToString(), "application/json" }
              },
                Content = new StringContent(bodyTokenMilvus, Encoding.UTF8, "application/json")
            };

            HttpResponseMessage response2 = await client.SendAsync(httpRequestToken);

            if (response2.StatusCode == HttpStatusCode.Found)
            {
                string responseBody = await response2.Content.ReadAsStringAsync();
                string jwt2 = JsonDocument.Parse(responseBody).RootElement.GetProperty("jwt2").GetString();

                string urlMilvus = "https://api.uat.errepar.com/syserrepar/embeddeddocumentai/loadAssetToMilvus?useSearchTermsGeneration=" + isIA;

                var httpRequestMilvus = new HttpRequestMessage
                {
                    Method = HttpMethod.Post,
                    RequestUri = new Uri(urlMilvus),
                    Headers = {
                      { HttpRequestHeader.ContentType.ToString(), "application/json" },
                      { HttpRequestHeader.Authorization.ToString(), jwt2 }
                  },
                    Content = new StringContent(json, Encoding.UTF8, "application/json")
                };

                HttpResponseMessage responseMilvus = await client.SendAsync(httpRequestMilvus);

                if (responseMilvus.StatusCode == HttpStatusCode.OK)
                {
                    Console.WriteLine($"Activo {id} migrado a MILVUS");
                    return (true, "200");
                }
                else
                {
                    //GrabarReporte(String.Format("Activo {0} fallo MILVUS| Hora: {1} - {2} | Error: {3}", id, DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString(), response2.ReasonPhrase), "C:\\Users\\gonzalo.sanchez\\source\\repos\\Errepar.Alpha.MigradorMasivo\\Logs\\DocumentosErrorLegisRedirect-" + _nombreBiblioteca + _ambiente + "2.txt");
                    //Console.WriteLine("milvus - ERROR en activo " + id);
                    var reason = responseMilvus.ReasonPhrase ?? "";
                    Console.WriteLine("milvus - ERROR en activo " + id + " | " + reason);
                    
                    var responseBodyError = await responseMilvus.Content.ReadAsStringAsync();
                    Console.WriteLine($"MILVUS - ERROR en activo {id} | {reason}");
                    Console.WriteLine($"Response body: {responseBodyError}");
                    return (false, $"{(int)responseMilvus.StatusCode} {reason}");
                }
            }
            else
            {
                //GrabarReporte(String.Format("Activo {0} fallo TOKEN MILVUS| Hora: {1} - {2} | Error: {3}", id, DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString(), response2.ReasonPhrase), "C:\\Users\\gonzalo.sanchez\\source\\repos\\Errepar.Alpha.MigradorMasivo\\Logs\\DocumentosErrorLegisRedirect-" + _nombreBiblioteca + _ambiente + "2.txt");
                //Console.WriteLine("ERROR ");
                var reason = response2.ReasonPhrase ?? "";
                Console.WriteLine("ERROR TOKEN MILVUS | " + reason);
                Console.WriteLine(response2);
                Console.WriteLine($"ERROR TOKEN MILVUS | {reason}");
                return (false, $"Token {(int)response2.StatusCode} {reason}");
            }
        }

    }
}