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

namespace Errepar.MetadataManager.Process.Services
    {
    // Esqueleto de cliente para Solr. Ajusta URL y esquema de campos según tu core/index.
    public class SolrManager
        {
        private readonly HttpClient _http;
        private readonly string _solrBaseUrl; // p. ej. http://localhost:8983/solr/yourcore;
        private readonly string _ambiente = "UAT";
        public SolrManager(HttpClient httpClient, string solrBaseUrl, string ambiente)
            {
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _solrBaseUrl = solrBaseUrl?.TrimEnd('/') ?? throw new ArgumentNullException(nameof(solrBaseUrl));
            _ambiente = ambiente.ToUpperInvariant();
            }

        public async Task<(bool ok, string msg)> AtomicUpdateSolrMultiple(string guid, Dictionary<string, object> campos, int idElemento)
            {
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


            return await EnviarASolr(jsonPayload, idElemento, _ambiente);
            }
        private async Task<(bool ok, string msg)> EnviarASolr(string json, int id, string ambiente)
            {
            //string user = "Basic " +  Base64Encode("uat-solr-acess:SXtuCde9&JW%pkAK");

            string user = "Basic " + (ambiente.Equals("PROD") ? Base64Encode("admin:6s2HUXFb8la") : Base64Encode("uat-solr-acess:SXtuCde9&JW%pkAK"));
            //string urlSolr = ambiente.Equals("PROD") ? "https://solr-prod-hcs.errepar.com/solr/prodActivos02/update/json/docs?commit=true" : "https://solr.errepar.com/solr/prodActivos01/update/json/docs?commit=true";
            string urlSolr = "https://solr.uat.errepar.com/solr/prodActivos02/update?commit=true";
            //string urlSolr = "https://solr-prod-hcs.errepar.com/solr/prodActivos02/update/json/docs?commit=true";

            //string urlSolr = "https://solr.errepar.com/solr/prodActivos02/update?commit=true";

            var httpRequestMessage = new HttpRequestMessage
                {
                Method = HttpMethod.Post,
                RequestUri = new Uri(urlSolr),
                Headers = {
             { HttpRequestHeader.ContentType.ToString(), "application/json" },
             { HttpRequestHeader.Authorization.ToString(), user }
         },
                Content = new StringContent(json, Encoding.UTF8, "application/json")
                };

            var response2 = _http.SendAsync(httpRequestMessage).Result;

            if (response2.StatusCode == HttpStatusCode.OK)
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
                var reason = response2.ReasonPhrase ?? "";
                Console.WriteLine("SOLR - ERROR en activo " + id + " | " + reason);
                return (false, $"{(int)response2.StatusCode} {reason}");
                }
            }

        public static string Base64Encode(string plainText)
            {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
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
                if ((field.InternalName.StartsWith("eolShp") || field.InternalName.StartsWith("Sumario_x0020_")))
                    {
                    if (field.TypeDisplayName.ToLower().Equals("metadatos administrados"))
                        {
                        if (field.TypeAsString.Equals("TaxonomyFieldTypeMulti"))
                            {
                            List<string> values = new List<string>();
                            List<string> valuesGUID = new List<string>();
                            TaxonomyFieldValueCollection _taxonomyValueColl = (item[field.InternalName] as TaxonomyFieldValueCollection);
                            foreach (var _taxonomyValue in _taxonomyValueColl)
                                {

                                values.Add(_taxonomyValue.Label);

                                if (field.InternalName.Equals("eolShpIndiceContenidosEOL"))
                                    {
                                    string guid = await getFullPathIds(_taxonomyValue.TermGuid, "Índice de Contenidos");
                                    valuesGUID.Add(guid);
                                    }
                                else if (field.InternalName.Equals("eolShpIndiceContenidosIUS"))
                                    {
                                    string guid = await getFullPathIds(_taxonomyValue.TermGuid, "Índice de Contenidos IUS");
                                    valuesGUID.Add(guid);
                                    }

                                }

                            if (field.InternalName.Equals("eolShpIndiceContenidosEOL"))
                                {
                                col["eolShpIndiceContenidosGUID"] = valuesGUID.ToArray();
                                }
                            else if (field.InternalName.Equals("eolShpIndiceContenidosIUS"))
                                {
                                col["eolShpIndiceContenidosIUSGUID"] = valuesGUID.ToArray();
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
                        if (field.InternalName.StartsWith("Sumario_x0020_"))
                            {
                            col[field.InternalName.Replace("Sumario_x0020_", "eolShpResumen")] = item[field.InternalName];
                            }
                        else
                            col[field.InternalName] = item[field.InternalName];
                        }

                    }
                }

            col["eolShpID"] = item.Id;
            col["ModerationStatus"] = "Approved";
            //col["searchable"] = "1";
            col["eolShpTipoContenido"] = item.ContentType.Name;


            ctx.Load(item.File);
            var binaryStream = item.File.OpenBinaryStream();
            ctx.ExecuteQuery();

            using (var sr = new System.IO.StreamReader(binaryStream.Value))
                {
                var line = sr.ReadToEnd();
                col["eolShpBody"] = line;
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
