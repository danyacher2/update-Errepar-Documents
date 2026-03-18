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

        public static async Task<(bool ok, string msg)> EnviarAMilvus(HttpClient client,string json, int id,bool isIA = false)
        {
            //string bodyTokenMilvus = "{\"key\":\"e5169fe4-1c39-4356-b148-1f210cc2431b\"}";
            //string urlTokenMilvus = "https://accounts.errepar.com/eauth/auth/job/getJobCredentialsEAuth";


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
                //Console.WriteLine(GetResponseBodyAsync(response2));
                //string urlMilvus = "https://api.errepar.com/syserrepar/embeddeddocumentai/loadAssetToMilvus?useSearchTermsGeneration=false";
                string urlMilvus = "https://api.uat.errepar.com/syserrepar/embeddeddocumentai/loadAssetToMilvus?useSearchTermsGeneration="+isIA;

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
                    //GrabarReporte(String.Format("Activo {0} migrado | Hora: {1} - {2}", id, DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString()), "C:\\Users\\gonzalo.sanchez\\source\\repos\\Errepar.Alpha.MigradorMasivo\\Logs\\DocumentosProcesadosLegisRedirect-" + _nombreBiblioteca + _ambiente + "2.txt");
                    Console.WriteLine("Activo MIGRADO MILVUS");
                    return (true, "200");
                }
                else
                {
                    //GrabarReporte(String.Format("Activo {0} fallo MILVUS| Hora: {1} - {2} | Error: {3}", id, DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString(), response2.ReasonPhrase), "C:\\Users\\gonzalo.sanchez\\source\\repos\\Errepar.Alpha.MigradorMasivo\\Logs\\DocumentosErrorLegisRedirect-" + _nombreBiblioteca + _ambiente + "2.txt");
                    //Console.WriteLine("milvus - ERROR en activo " + id);
                    var reason = responseMilvus.ReasonPhrase ?? "";
                    Console.WriteLine("milvus - ERROR en activo " + id + " | " + reason);
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
                return (false, $"Token {(int)response2.StatusCode} {reason}");
            }
        }

    }
}