using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;
using Errepar.MetadataManager.Process.Config;

namespace Errepar.MetadataManager.Process.Services
{
    // Esqueleto de cliente para Solr. Ajusta URL y esquema de campos según tu core/index.
    public class SolrManager
    {
        private readonly HttpClient _http;
        private readonly SolrSettings _solrSettings;
        private readonly MilvusSettings _milvusSettings;

        public SolrManager(HttpClient httpClient, SolrSettings solrSettings, MilvusSettings milvusSettings)
        {
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _solrSettings = solrSettings ?? throw new ArgumentNullException(nameof(solrSettings));
            _milvusSettings = milvusSettings ?? throw new ArgumentNullException(nameof(milvusSettings));
        }

        public async Task<(bool ok, string msg)> AtomicUpdateSolrMultiple(string guid, Dictionary<string, object> campos, int idElemento)
        {
            const int maxIntentos = 3;
            const int delayEntreIntentos = 3000; // 2 segundos

            // Armamos el bloque dinámico para cada campo con su tipo correcto
            var updates = string.Join(",\n", campos.Select(kv =>
            {
                string valorFormateado;

                switch (kv.Value)
                {
                    case bool b:
                        valorFormateado = b.ToString().ToLowerInvariant();
                        break;
                    case string s:
                        valorFormateado = $"\"{s}\"";
                        break;
                    case DateTime dt:
                        valorFormateado = $"\"{dt.ToUniversalTime():yyyy-MM-ddTHH:mm:ssZ}\"";
                        break;
                    case IEnumerable<string> lista:
                        valorFormateado = "[ " + string.Join(", ", lista.Select(v => $"\"{v}\"")) + " ]";
                        break;
                    default:
                        valorFormateado = $"\"{kv.Value}\"";
                        break;
                }

                return $@"""{kv.Key}"": {{ ""set"": {valorFormateado} }}";
            }));

            string jsonPayload = $@"
         [
             {{
                 ""id"": ""{guid}"",
                 {updates}
             }}
         ]";

            // Política de reintentos
            int intentoActual = 0;
            (bool ok, string msg) resultado = (false, string.Empty);

            while (intentoActual < maxIntentos)
            {
                intentoActual++;

                try
                {
                    Console.WriteLine($"🔄 SolrManager - Intento {intentoActual}/{maxIntentos} para elemento {idElemento}");

                    resultado = await EnviarASolr(jsonPayload, idElemento);
                    await Task.Delay(30);

                    if (resultado.ok)
                    {
                        if (intentoActual > 1)
                        {
                            Console.WriteLine($"✔️ Solr exitoso en intento {intentoActual}/{maxIntentos} para elemento {idElemento}");
                        }
                        return resultado;
                    }

                    Console.WriteLine($"⚠️ SolrManager - Error en intento {intentoActual}/{maxIntentos}: {resultado.msg}");

                    if (intentoActual < maxIntentos)
                    {
                        await Task.Delay(delayEntreIntentos);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ SolrManager - Excepción en intento {intentoActual}/{maxIntentos}: {ex.Message}");
                    Console.WriteLine($"⚠️ SolrManager - Excepción en intento {intentoActual}/{maxIntentos}: {ex}");

                    if (intentoActual >= maxIntentos)
                    {
                        return (false, $"Excepción después de {maxIntentos} intentos: {ex.Message}");
                    }

                    await Task.Delay(delayEntreIntentos);
                }
            }
            Console.WriteLine($" {resultado}");

            return (false, $"Falló después de {maxIntentos} intentos. Último error: {resultado.msg}");
        }

        public async Task<(bool ok, string msg)> EnviarASolr(string json, int id)
        {
            string user = "Basic " + Base64Encode($"{_solrSettings.User}:{_solrSettings.Password}");
            string urlSolr = _solrSettings.UpdateUrl;

            // Crear un NUEVO HttpRequestMessage en cada llamada (no reutilizar)
            using var httpRequestMessage = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(urlSolr),
                Content = new StringContent(json, Encoding.UTF8, "application/json"),
            };

            httpRequestMessage.Headers.Add("Authorization", user);

            // IMPORTANTE: Usar await en lugar de .Result para evitar deadlocks
            using var response = await _http.SendAsync(httpRequestMessage);

            if (response.StatusCode == HttpStatusCode.OK)
            {
                //GrabarReporte(String.Format("Activo {0} migrado | Hora: {1} - {2}", id, DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString()), "C:\\Users\\gonzalo.sanchez\\source\\repos\\Errepar.Alpha.MigradorMasivo\\Logs\\DocumentosProcesados-" + _nombreBiblioteca + _ambiente + ".txt");
                Console.WriteLine("Activo " + id + " migrado a SOLR");
                return (true, "200");
                //Console.WriteLine("Enviando a Milvus..." + json);

                //  EnviarAMilvus(json, id);
            }
            else
            {

                //GrabarReporte(String.Format("Activo {0} fallo SOLR| Hora: {1} - {2} | Error: {3}", id, DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString(), response2.ReasonPhrase), "C:\\Users\\gonzalo.sanchez\\source\\repos\\Errepar.Alpha.MigradorMasivo\\Logs\\DocumentosError-" + _nombreBiblioteca + _ambiente + "2.txt");
                //Console.WriteLine("SOLR - ERROR en activo " + id);
                var reason = response.ReasonPhrase ?? "";
                var responseBody = await response.Content.ReadAsStringAsync();
                Console.WriteLine("SOLR - ERROR en activo " + id + " | " + reason);
                Console.WriteLine($"Response body: {responseBody}");
                return (false, $"{(int)response.StatusCode} {reason}");
            }
        }

        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public async void EliminarDeSolr(string eolShpTimeStamp)
        {
            //string jsonData = "{\"delete\":{\"query\":\"eolShpTimeStamp:\"" + eolShpTimeStamp + "\"\"}}";
            // string jsonData = "{'delete':{'eolShpTimeStamp:" + eolShpTimeStamp + "'}}";

            //string jsonData = @"{
            //      ""delete"": {
            //        ""query"": ""eolShpTimeStamp:" + eolShpTimeStamp+@"""
            //      },
            //      ""commit"": {}
            //    }
            //    ";

            //string jsonData = "{\"delete\":{\"query\":\"eolShpTimeStamp:"+ eolShpTimeStamp + "\"}}";

            var payload = new
            {
                delete = new
                {
                    query = $"eolShpTimeStamp:\"{eolShpTimeStamp}\""
                }
            };

            string json = JsonConvert.SerializeObject(payload);
            // string json = JsonConvert.SerializeObject(jsonData);

            string user = "Basic " + Base64Encode($"{_solrSettings.User}:{_solrSettings.Password}");
            string urlSolr = _solrSettings.UpdateUrl;
            HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");
            var httpRequestMessage = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(urlSolr),
                Headers = {
                    { HttpRequestHeader.ContentType.ToString(), "application/json" },
                    { HttpRequestHeader.Authorization.ToString(), user }
                },
                Content = content//new StringContent(json)
            };

            var response2 = this._http.SendAsync(httpRequestMessage).Result;

            if (response2.StatusCode == HttpStatusCode.OK)
            {
                Console.WriteLine("Activo " + eolShpTimeStamp + " eliminado");
                await EliminarDeMilvus(eolShpTimeStamp);
            }
            else
            {
                Console.WriteLine("ERROR en activo " + eolShpTimeStamp);
            }
        }

        private async Task<(bool ok, string msg)> EliminarDeMilvus(string eolShpTimeStamp)
        {
            string bodyTokenMilvus = "{\"key\":\"" + _milvusSettings.ApiKey + "\"}";
            string urlTokenMilvus = _milvusSettings.AuthEndpoint;

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

            HttpResponseMessage response2 = this._http.SendAsync(httpRequestToken).Result;

            if (response2.StatusCode == HttpStatusCode.Found)
            {
                string responseBody = await response2.Content.ReadAsStringAsync();
                string jwt2 = JsonDocument.Parse(responseBody).RootElement.GetProperty("jwt2").GetString();
                //Console.WriteLine(GetResponseBodyAsync(response2));
                string urlMilvus = $"{_milvusSettings.DeleteAssetEndpoint}?eolShpTimeStamp={eolShpTimeStamp}";

                var httpRequestMilvus = new HttpRequestMessage
                {
                    Method = HttpMethod.Delete,
                    RequestUri = new Uri(urlMilvus),
                    Headers = {
                        { HttpRequestHeader.ContentType.ToString(), "application/json" },
                        { HttpRequestHeader.Authorization.ToString(), jwt2 }
                    },
                    //Content = new StringContent(json, Encoding.UTF8, "application/json")
                };

                HttpResponseMessage responseMilvus = this._http.SendAsync(httpRequestMilvus).Result;

                if (responseMilvus.StatusCode == HttpStatusCode.OK)
                {
                    //GrabarReporte(String.Format("Activo {0} eliminado milvus", id), "C:\\Users\\gonzalo.sanchez\\source\\repos\\Errepar.Alpha.MigradorMasivo\\Logs\\Duplicados-Eliminados" + _nombreBiblioteca + _ambiente + ".txt");
                    Console.WriteLine("Activo " + eolShpTimeStamp + " eliminado milvus");
                    return (true, "200");
                }
                else
                {
                    var reason = responseMilvus.ReasonPhrase ?? "";
                    Console.WriteLine("milvus - ERROR en activo " + eolShpTimeStamp + " | " + reason);
                    return (false, $"{(int)responseMilvus.StatusCode} {reason}");
                }
            }
            else
            {
                var reason = response2.ReasonPhrase ?? "";
                Console.WriteLine("ERROR TOKEN MILVUS | " + reason);
                return (false, $"Token {(int)response2.StatusCode} {reason}");
            }
        }


        public async Task<IDictionary<string, object>> GetItemFormatSync(ClientContext ctx, ListItem item)
        {
            IDictionary<string, object> col = new Dictionary<string, object>();
            col["id"] = item["GUID"];
            col["eolShpTipoContenido"] = item.ContentType.Name;
            //col["eolShpTipoContenido"] = _tipoDeContenido;
            col["eolShpTitle"] = item["Title"];

            foreach (Field field in item.ParentList.Fields)
            {
                //if((field.InternalName.StartsWith("eolShp") || field.InternalName.StartsWith("Sumario_x0020_")) && !field.InternalName.Equals("eolShpMarcasCD") && !field.InternalName.Equals("eolShpIndiceContenidosIUS") && !field.InternalName.Equals("eolShpIndiceContenidosOculto") && item[field.InternalName] != null) 
                if (field.InternalName.StartsWith("eolShp") || field.InternalName.StartsWith("Sumario_x0020_"))
                {
                    if (field.TypeDisplayName.ToLower().Equals("metadatos administrados"))
                    {
                        if (field.TypeAsString.Equals("TaxonomyFieldTypeMulti"))
                        {
                            List<string> values = new List<string>();
                            List<string> valuesGUID = new List<string>();
                            var rawTaxonomyValue = item[field.InternalName];
                            var taxonomyTerms = new List<(string Label, string TermGuid)>();

                            if (rawTaxonomyValue is TaxonomyFieldValueCollection taxonomyValueCollection2)
                            {
                                foreach (var value in taxonomyValueCollection2)
                                    taxonomyTerms.Add((value.Label, value.TermGuid));
                            }
                            else if (rawTaxonomyValue is string serializedTaxonomyValue)
                            {
                                var taxonomyField = ctx.CastTo<TaxonomyField>(field);
                                var taxonomyValueCollection = new TaxonomyFieldValueCollection(
                                    ctx,
                                    serializedTaxonomyValue,
                                    taxonomyField);
                                foreach (var value in taxonomyValueCollection)
                                    taxonomyTerms.Add((value.Label, value.TermGuid));
                            }
                            else if (rawTaxonomyValue != null)
                            {
                                var childItems = Newtonsoft.Json.Linq.JObject.FromObject(rawTaxonomyValue)["_Child_Items_"]
                                    as Newtonsoft.Json.Linq.JArray;
                                if (childItems != null)
                                {
                                    foreach (var childItem in childItems)
                                    {
                                        var label = (string?)childItem["Label"];
                                        var termGuid = (string?)childItem["TermGuid"];
                                        if (label != null && termGuid != null)
                                            taxonomyTerms.Add((label, termGuid));
                                    }
                                }
                            }

                            foreach (var taxonomyTerm in taxonomyTerms)
                            {
                                values.Add(taxonomyTerm.Label);

                                if (field.InternalName.Equals("eolShpIndiceContenidosEOL"))
                                {
                                    string guid = await getFullPathIds(taxonomyTerm.TermGuid, "Índice de Contenidos");
                                    valuesGUID.Add(guid);
                                }
                                else if (field.InternalName.Equals("eolShpIndiceContenidosIUS"))
                                {
                                    string guid = await getFullPathIds(taxonomyTerm.TermGuid, "Índice de Contenidos IUS");
                                    valuesGUID.Add(guid);
                                }

                            }

                            if (field.InternalName.Equals("eolShpIndiceContenidosEOL"))
                            {
                                col["eolShpIndiceContenidosGUID"] = valuesGUID.ToArray();
                                col["eolShpIndiceContenidosGUID_entries"] = valuesGUID.ToArray();
                            }
                            else if (field.InternalName.Equals("eolShpIndiceContenidosIUS"))
                            {
                                col["eolShpIndiceContenidosIUSGUID"] = valuesGUID.ToArray();
                                col["eolShpIndiceContenidosIUSGUID_entries"] = valuesGUID.ToArray();
                            }

                            if (field.InternalName.Equals("eolShpOrganismo"))
                                col[field.InternalName] = NormalizeOrganismo(values.ToArray());
                            else
                                col[field.InternalName] = values.ToArray();
                        }
                        else
                        {
                            TaxonomyFieldValue _taxonomyValue = item[field.InternalName] as TaxonomyFieldValue;
                            if (_taxonomyValue != null)
                                col[field.InternalName] = _taxonomyValue.Label;
                        }
                    }
                    else if (field.TypeDisplayName.ToLower().Equals("búsqueda"))
                    {
                        if (field.TypeAsString.Equals("LookupMulti"))
                        {
                            List<string> values = new List<string>();
                            FieldLookupValue[] _lookupValueArr = item[field.InternalName] as FieldLookupValue[];
                            if (_lookupValueArr != null)
                            {
                                foreach (var _lookupValue in _lookupValueArr)
                                {
                                    if (_lookupValue != null && !string.IsNullOrEmpty(_lookupValue.LookupValue))
                                        values.Add(_lookupValue.LookupValue);
                                }
                            }
                            if (values.Count > 0)
                                col[field.InternalName] = values.ToArray();
                        }
                        else
                        {
                            FieldLookupValue _lookupValue = item[field.InternalName] as FieldLookupValue;

                            if (_lookupValue != null)
                                col[field.InternalName] = _lookupValue.LookupValue;
                        }
                    }
                    else if (field.TypeDisplayName.ToLower().Equals("número")) //DateTime
                    {
                        try
                        {
                            if (field.InternalName.Equals("eolShpOrden"))
                                col[field.InternalName] = Convert.ToInt32(item[field.InternalName]);
                            else
                                col[field.InternalName] = Convert.ToInt32(item[field.InternalName]);
                        }
                        catch (Exception)
                        {
                            col[field.InternalName] = item[field.InternalName];
                        }
                    }
                    else if (field.TypeDisplayName.ToLower().Equals("fecha y hora"))
                    {
                        if (item.ContentType.Name.Equals("Legislacion"))
                        {
                            if (field.InternalName.Equals("eolShpFecha") && item[field.InternalName] != string.Empty)
                                col["eolShpFechaBusqueda"] = item[field.InternalName];

                            if (!field.InternalName.Equals("eolShpFechaBusqueda"))
                                col[field.InternalName] = item[field.InternalName];
                        }
                        else
                            col[field.InternalName] = item[field.InternalName];
                    }
                    else
                    {
                        var fieldValue = item[field.InternalName];
                        if (fieldValue is FieldUrlValue urlValue)
                            fieldValue = urlValue.Url;

                        if (field.InternalName.StartsWith("Sumario_x0020_"))
                        {
                            col[field.InternalName.Replace("Sumario_x0020_", "eolShpResumen")] = fieldValue;
                        }
                        else
                            col[field.InternalName] = fieldValue;
                    }

                }
            }

            col["eolShpID"] = item.Id;
            //col["idd"] = "b05491dc-08b5-41b6-84b3-3d99701fb381";
            col["ModerationStatus"] = "Approved";
            //col["searchable"] = "1";
            //col["eolShpTipoContenido"] = item.ContentType.Name;

            if (item.File != null)
            {
                ctx.Load(item.File);
                var binaryStream = item.File.OpenBinaryStream();
                ctx.ExecuteQuery();

                using (var sr = new System.IO.StreamReader(binaryStream.Value))
                {
                    var line = sr.ReadToEnd();
                    col["eolShpBody"] = line;
                }
            }

            return col;
        }
        //public async Task IndexAsync(IEnumerable<object> documents, CancellationToken cancellationToken = default)
        //{
        //    // POST a /update?commit=true con JSON de documentos
        //    var url = $"{_solrBaseUrl}/update?commit=true";
        //    await _http.PostAsJsonAsync(url, documents, cancellationToken).ConfigureAwait(false);
        //}

        //public async Task DeleteByIdAsync(IEnumerable<string> ids, CancellationToken cancellationToken = default)
        //{
        //    var url = $"{_solrBaseUrl}/update?commit=true";
        //    var deletePayload = new { delete = ids };
        //    await _http.PostAsJsonAsync(url, deletePayload, cancellationToken).ConfigureAwait(false);
        //}

        //public async Task<string> SearchRawAsync(string q, int rows = 10, CancellationToken cancellationToken = default)
        //{
        //    var url = $"{_solrBaseUrl}/select?q={Uri.EscapeDataString(q)}&rows={rows}&wt=json";
        //    var resp = await _http.GetStringAsync(url, cancellationToken).ConfigureAwait(false);
        //    return resp;
        //}
        private static string[] NormalizeOrganismo(string[] values)
        {
            if (values == null || values.Length == 0)
                return new string[0];

            var list = new List<string>();

            foreach (var raw in values)
            {
                if (string.IsNullOrWhiteSpace(raw))
                    continue;

                string last = raw;
                int idx = raw.LastIndexOf(':');
                if (idx >= 0 && idx < raw.Length - 1)
                {
                    last = raw.Substring(idx + 1);
                }

                string cleaned = last.Trim();
                if (!string.IsNullOrWhiteSpace(cleaned))
                {
                    list.Add(cleaned);
                }
            }

            return list.ToArray();
        }

        public static async Task<string> getFullPathIds(string TermId, String TermSet)
        {
            try
            {
                HttpClient newClient = new HttpClient();
                var urlRequest = string.Format("https://getfullpathids.azurewebsites.net/api/calculateFullPathIds?termId={0}&termset={1}", TermId, TermSet);
                HttpRequestMessage newRequest = new HttpRequestMessage(HttpMethod.Get, urlRequest);

                //Read Server Response
                HttpResponseMessage res = await newClient.SendAsync(newRequest);
                string isValidMpn = await res.Content.ReadAsStringAsync();
                //Console.WriteLine($"Contenido de la respuesta:\n{isValidMpn}");
                return isValidMpn;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
            }
        }

        public static async Task<string> GetResponseBodyAsync(HttpResponseMessage response)
        {
            string responseBody = await response.Content.ReadAsStringAsync();
            return responseBody;
        }
    }



}
