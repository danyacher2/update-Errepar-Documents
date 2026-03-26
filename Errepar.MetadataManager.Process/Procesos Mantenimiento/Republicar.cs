using Errepar.MetadataManager.Process.Auth;
using Errepar.MetadataManager.Process.Services;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;

namespace Errepar.MetadataManager.Process.Procesos_Mantenimiento
{
    internal class Republicar
    {
        public async Task SoleryMulvus()
        {
            Console.WriteLine("Consulta a SharePoint: recuperar items con EstadoProceso = Pendiente | En Pausa");

            string siteUrl = "https://erreparsa.sharepoint.com/sites/ErreparDesarrollo";
            string tenantId = "00f26ad1-2073-4746-a79f-c83061db35c0";
            //string clientId = "8688eed4-7464-4288-9820-34849fd19296";//erreparDesarrollo
            string clientId = "f679c472-7b0c-45dc-b38c-cca0b662f77a";//erreparDev

            string certificateThumbprint = Environment.GetEnvironmentVariable("CERT_THUMBPRINT")
                ?? "452079A2697BC9646023FAE02876488654BBDB2C";
            var authToken = await TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync(tenantId, clientId, certificateThumbprint, siteUrl);

            // Configurar HttpClientHandler para aceptar certificados SSL y manejar problemas DNS (necesario para UAT/VPN)
            var socketHandler = new SocketsHttpHandler
            {
                SslOptions = new System.Net.Security.SslClientAuthenticationOptions
                {
                    RemoteCertificateValidationCallback = (sender, cert, chain, sslPolicyErrors) => true,
                    EnabledSslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13,
                    CertificateRevocationCheckMode = System.Security.Cryptography.X509Certificates.X509RevocationMode.NoCheck
                },
                UseProxy = false,
                PooledConnectionLifetime = TimeSpan.FromMinutes(5),
                MaxConnectionsPerServer = 20,
                ConnectTimeout = TimeSpan.FromSeconds(30)
            };

            using var http = new HttpClient(socketHandler);
            http.Timeout = TimeSpan.FromMinutes(10);

            var context = new ClientContext(siteUrl);

            var solerManager = new SolrManager(http, siteUrl, "UAT");

            context.ExecutingWebRequest += (sender, e) =>
            {
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authToken;
            };



            var libraryNames = new List<string>
{
    //"Documento",
    //"Agenda",
    //"Doctrina",
    //"Guía Temática",
    //"Jurisprudencia Adm",
    //"Jurisprudencia Judicial",
    //"Legislación",
    //"Modelo",
    "Actualidad",
    "Videos",
    "Podcast"
};


            foreach (var name in libraryNames)
            {
                var list = context.Web.Lists.GetByTitle(name);
                context.Load(list);

                Console.WriteLine($"Lista: {name}");

            }
            List<string> itemsGuid = new List<string>();
            foreach (var listName in libraryNames)
            {
                try
                {
                    Console.WriteLine($"Procesando lista/biblioteca: {listName}");

                    List list = context.Web.Lists.GetByTitle(listName);
                    context.Load(list);
                    context.ExecuteQuery();

                    CamlQuery query = new CamlQuery
                    {
                        ViewXml = @"
<View Scope='RecursiveAll'>
    <RowLimit Paged='TRUE'>2000</RowLimit>
</View>"
                    };

                    ListItemCollectionPosition position = null;

                    do
                    {
                        query.ListItemCollectionPosition = position;

                        ListItemCollection itemsList = list.GetItems(query);

                        context.Load(itemsList);
                        context.ExecuteQuery();

                        foreach (ListItem item in itemsList)
                        {
                            context.Load(item, i => i.ContentType, i => i.ParentList, i => i.ParentList.Fields, i => i["eolShpTema"]);
                            context.ExecuteQuery();

                            if (string.Equals(item.ContentType?.Name, "carpeta", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            var fileName = item["FileLeafRef"]?.ToString();

                            if (fileName.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                                fileName.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                                fileName.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) ||
                                fileName.EndsWith(".gif", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }



                            Guid itemGuid = Guid.Empty;
                            if (item.FieldValues.ContainsKey("GUID") && item["GUID"] != null)
                            {
                                itemGuid = (Guid)item["GUID"];
                            }

                            int id = 0;
                            if (item.FieldValues.ContainsKey("ID") && item["ID"] != null)
                            {
                                id = Convert.ToInt32(item["ID"]);
                            }
                            Console.WriteLine($"{itemGuid}");
                            itemsGuid.Add(itemGuid.ToString());

                            try
                            {
                                var timeStamp = item["eolShpTimeStamp"].ToString();
                                if (timeStamp != null && timeStamp != "")
                                {
                                    //elimino si existe en soler o milvus
                                    solerManager.EliminarDeSolr(timeStamp, "UAT");
                                    await Task.Delay(30);
                                }
                            }
                            catch (Exception ex)
                            {
                                System.IO.File.AppendAllLines("errorsToDelete.txt", [itemGuid.ToString()], Encoding.UTF8);

                            }


                            //enviar a soler:

                            var itemFormateado = await solerManager.GetItemFormatSync(context, item);

                            // Armamos el bloque dinámico para cada campo con su tipo correcto
                            var updates = string.Join(",\n", itemFormateado.Select(kv =>
                            {
                                string valorFormateado = "";
                                if (kv.Key == "id")
                                    return $@"""{kv.Key}"":""{kv.Value}""";
                                if (kv.Key == "eolShpBody")
                                    return $@"""{kv.Key}"":""{kv.Value.ToString().Replace('"', '\'')}""";
                                switch (kv.Value)
                                {
                                    case bool b:
                                        return $@"""{kv.Key}"":{b.ToString().ToLower()}";
                                    case string s:
                                        valorFormateado = $"\"{s}\"";
                                        break;
                                    case DateTime dt:
                                        valorFormateado = $"\"{dt.ToUniversalTime():yyyy-MM-ddTHH:mm:ssZ}\"";
                                        break;
                                    case IEnumerable<string> lista:
                                        valorFormateado = "[" + string.Join(", ", lista.Select(v => $"\"{v}\"")) + "]";
                                        break;

                                    default:
                                        if (kv.Value != null)
                                            return $@"""{kv.Key}"":{kv.Value}";
                                        return $@"""{kv.Key}"":""{kv.Value}""";

                                        break;
                                }
                                return $@"""{kv.Key}"":{valorFormateado}";

                            }));
                            string jsonPayload = $@"
                     [
                         {{
                             {updates}
                         }}
                     ]";
                            //Console.WriteLine(jsonPayload);
                            var resultado = await solerManager.EnviarASolr(jsonPayload, item.Id, "UAT");
                            await Task.Delay(30);

                            if (resultado.ok)
                            {

                                Console.WriteLine($"✔️ Se agrego/modifico el elemento a soler: {itemGuid}");

                            }
                            else
                            {
                                Console.WriteLine($"⚠️ Error en soler: {itemGuid}: {resultado}");

                            }

                            string json = JsonConvert.SerializeObject(itemFormateado, Newtonsoft.Json.Formatting.Indented);
                            // Enviar a Milvus
                            var milvusResult = await MilvusManager.EnviarAMilvus(http, json, item.Id, false);

                            if (milvusResult.ok)
                            {
                                Console.WriteLine($"✔️ Milvus actualizado para ID {itemGuid}");
                            }
                            else
                            {
                                Console.WriteLine($"⚠️ Error al actualizar en Milvus: {milvusResult.msg}");
                            }
                        }


                        position = itemsList.ListItemCollectionPosition;
                    }
                    while (position != null);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"No se pudo procesar '{listName}': {ex.Message}");
                }
            }
            System.IO.File.WriteAllLines("ids3.txt", itemsGuid, Encoding.UTF8);

            return;
        }
    }
}
