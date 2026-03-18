using Errepar.MetadataManager.Process.Auth;
using Errepar.MetadataManager.Process.Models;
using Errepar.MetadataManager.Process.Services;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.News.DataModel;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Security;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using static System.Net.WebRequestMethods;

var jsonOptions = new JsonSerializerOptions
{
    WriteIndented = true
};


try
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

    var logsManager = new LogsManager(http, siteUrl, "Metadata Manager");

    await logsManager.LogEjecucionAsync(0, "Inicio del proceso Metadata Manager", "Inicio");

    var searchService = new SearchService(context);

    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);

    var sp = new SharePointManager(http, siteUrl, "Metadata Manager", authToken);
    using var cts = new CancellationTokenSource(TimeSpan.FromHours(10));

    var items = await sp.GetPendingOrPausedAsync(cts.Token).ConfigureAwait(false);
    await logsManager.LogEjecucionAsync(0, $"Items recuperados desde SharePoint: {items.Count}", "OK");

    var cambios = new ProcesarCambios(http, context, siteUrl, logsManager, searchService);

    context.Load(context.Web, w => w.ServerRelativeUrl);
    context.ExecuteQuery();

    var webServerRelative = context.Web.ServerRelativeUrl;
    Console.WriteLine("ServerRelativeUrl del sitio: " + webServerRelative);

    var (allow, deny, rulesCount) = ScheduleGate.LoadWindows(context, "Manage Metadata Schedule");

    await logsManager.LogEjecucionAsync(0, $"Reglas de Schedule cargadas: {rulesCount}", "OK");

    var EstadoProceso = "Procesando";

    var (listas, listaProbable) = CargarBibliotecasSharePoint(context);
    var activosEncontrados = new List<string>();

    Console.WriteLine($"Items obtenidos: {items.Count}");

    for (int j = 0; j < items.Count; j++)
    {
        int CantidadProcesdos = 0;
        var it = items[j];
        Console.WriteLine($"Id={it.Id} | Título='{it.Titulo}' | Estado='{it.EstadoProceso}' | EjecutadoPor='{it.EjecutadoPor}' | Link='{it.Link}'");
        for (int i = 0; i < it.Activos.Count; i++)
        {
            var activo = it.Activos[i];

            var procesarEnShpSolrMilvus = it.Scope;


            Console.WriteLine($"{i} --Guid item: {activo}");

            EstadoProceso = await ManejarPausaPorSchedule(EstadoProceso, it.Id, allow, deny, sp, logsManager, cts.Token);

            var item = FindItem(listas, new Guid(activo), context, ref listaProbable);
            if (item == null)
                continue;

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

            Console.WriteLine($"{i} --Guid item: {activo}");

            Console.WriteLine($"item: {item}");
            if (item != null)
            {
                activosEncontrados.Add(activo);
            }
            if (item == null)
            {
                Console.WriteLine("❌ No se encontró el elemento en las listas candidatas");
                await logsManager.LogErrorAsync(it.Id, activo, $"Item con guid {activo} no se encontro", "No Encontrado");
                continue;
            }

            await logsManager.SaveItemOriginalJsonAsync(activo, item.FieldValues);

            if (it.Cambios != null)
            {
                var root = it.Cambios.RootElement;

                if (root.ValueKind == JsonValueKind.Array)
                {
                    foreach (var cambio in root.EnumerateArray())
                    {
                        cambios.Procesar(item, cambio, it.Id);
                    }
                }
                else if (root.ValueKind == JsonValueKind.Object)
                {
                    cambios.Procesar(item, root, it.Id);
                }
                item.SystemUpdate();
                context.ExecuteQuery();

                if (procesarEnShpSolrMilvus.Contains("Soler"))
                {
                    bool incluirMilvus = procesarEnShpSolrMilvus.Contains("Milvus");
                    await ProcesarEnSolr(it.Id, activo, item, root, solerManager, context, logsManager, incluirMilvus);
                }

                if (procesarEnShpSolrMilvus.Contains("Milvus - IA"))
                {
                    await ProcesarEnMilvusAsync(http, it.Id, activo, item.Id, item, root, solerManager, context, logsManager, true);
                }
                else if (procesarEnShpSolrMilvus.Contains("Milvus"))
                {
                    await ProcesarEnMilvusAsync(http, it.Id, activo, item.Id, item, root, solerManager, context, logsManager,false);
                }
            }
            CantidadProcesdos++;
            await logsManager.SaveItemProcessedJsonAsync(activo, item.FieldValues);

            await logsManager.LogEjecucionAsync(it.Id, "Inicio proceso", "OK");



            var logActivo = new ItemLogActivosProcesados
            {
                ItemCambiosId = it.Id,
                ItemListId = Convert.ToInt32(item.FieldValues["ID"]),
                LibraryName = item.FieldValues["Title"]?.ToString(),
                TimeStamp = Convert.ToInt32(item.FieldValues["ID"]),
                UrlItem = item.FieldValues["FileRef"]?.ToString(),
                Activo = activo,
                Procesado = true,
                Fecha = DateTime.Now
            };

            await logsManager.SaveLogActivosProcesados(it.Id, new[] { logActivo });
            //}
            //catch (Exception ex)
            //{

            await logsManager.LogErrorAsync(it.Id, activo, $"Item con guid {activo} no Modificado", "error en el cambio");
            //    Console.WriteLine("Error en la busqueda del Item"+ex);
            //}
            await sp.UpdateCantActivosAsync(it.Id, CantidadProcesdos);
        }


        await sp.UpdateEstadoProcesoAsync(it.Id, "Finalizado");


        await logsManager.SyncLogs(it.Id, cts.Token);

    }

    await logsManager.LogEjecucionAsync(0, "Proceso completado correctamente", "Finalizado");

}
catch (Exception ex)
{
    Console.Error.WriteLine("Excepción no controlada: " + ex);
    Environment.ExitCode = 1;
}
finally
{
    Console.WriteLine("Proceso finalizado con código " + Environment.ExitCode + ". Presiona una tecla para cerrar...");
    Console.ReadKey();
}

async Task<string> ManejarPausaPorSchedule(
    string estadoActual,
    int itemId,
    List<WeeklyWindow> allow,
    List<WeeklyWindow> deny,
    SharePointManager sp,
    LogsManager logsManager,
    CancellationToken cancellationToken)
{
    if (!ScheduleGate.IsAllowed(DateTimeOffset.Now, allow, deny))
    {
        if (estadoActual != "En Pausa")
        {
            estadoActual = "En Pausa";
            await sp.UpdateEstadoProcesoAsync(itemId, "En Pausa");
        }

        Console.WriteLine("Proceso en Pausa");
        await logsManager.LogEjecucionAsync(itemId, "Proceso en Pausa", "En Pausa");

        await ScheduleGate.WaitUntilAllowedAsync(allow, deny, TimeSpan.FromMinutes(1), cancellationToken);

        estadoActual = "Procesando";
        await sp.UpdateEstadoProcesoAsync(itemId, "Procesando");
        await logsManager.LogEjecucionAsync(itemId, "El Procesando", "Procesando");

        Console.WriteLine("Proceso Procesando");
    }

    return estadoActual;
}

(List<List>, List) CargarBibliotecasSharePoint(ClientContext context)
{
    var libraryNames = new List<string>
    {
        "Documento",
        "Agenda",
        "Doctrina",
        "Guía Temática",
        "Jurisprudencia Adm",
        "Jurisprudencia Judicial",
        "Legislación",
        "Modelo",
        "Actualidad",
        "Videos",
        "Podcast"
    };

    var listas = new List<List>();

    foreach (var name in libraryNames)
    {
        var list = context.Web.Lists.GetByTitle(name);
        context.Load(list);
        listas.Add(list);
    }

    var listaProbable = listas.FirstOrDefault();

    return (listas, listaProbable);
}

static ListItem GetItemByUniqueId(ClientContext ctx, List listaProbable, Guid uniqueId)
{
    try
    {
        var query = new CamlQuery
        {
            ViewXml = $@"
                        <View Scope='RecursiveAll'>
                            <Query>
                                <Where>
                                    <Eq>
                                        <FieldRef Name='GUID' />
                                        <Value Type='Guid'>{uniqueId}</Value>
                                    </Eq>
                                </Where>
                            </Query>
                            <RowLimit>1</RowLimit>
                        </View>"
        };

        var items = listaProbable.GetItems(query);
        ctx.Load(items);
        ctx.ExecuteQuery();

        if (items.Count > 0)
        {
            var item = items[0];

            //ctx.Load(item, i => i.ContentType, i => i.ParentList );
            try
            {
            ctx.Load(item, i => i.ContentType, i => i.ParentList, i => i.ParentList.Fields, i => i["eolShpTema"], i => i["eolShpCodigoMislibros"]);
            ctx.ExecuteQuery();
            }
            catch
            {
                ctx.Load(item, i => i.ContentType, i => i.ParentList, i => i.ParentList.Fields, i => i["eolShpTema"]);
                ctx.ExecuteQuery();
            }
            
            return item;
        }
    }
    catch(Exception Ex)
    {
        Console.WriteLine("Error"+Ex);
        // ignorar listas que fallen
    }


    return null;
}

string GetString(JsonElement el, string prop)
{
    if (el.TryGetProperty(prop, out var p))
        return p.ToString();

    return null;
}

ListItem FindItem(List<List> listas, Guid guid, ClientContext ctx, ref List listaProbable)
{

    foreach (var lista in listas)
    {
        var item = GetItemByUniqueId(ctx, lista, guid);

        if (item != null)
        {
            listaProbable = lista;
            return item;
        }
    }

    return null;
}

async Task<(IDictionary<string, object> itemFormateado, Dictionary<string, object> camposModificados)> ObtenerItemFormateadoYCampos(
    ListItem item,
    JsonElement root,
    SolrManager solrManager,
    ClientContext context)
{
    // 1️⃣ Formatear item
    var itemFormateado = await solrManager.GetItemFormatSync(context, item);

    // 2️⃣ Extraer campos modificados
    var campos = new Dictionary<string, object>();
    foreach (JsonElement cambio in root.EnumerateArray())
    {
        var campo = GetString(cambio, "campo");

        if (itemFormateado.TryGetValue(campo, out var valor))
        {
            if (!campos.ContainsKey(campo))
                campos.Add(campo, valor);
        }
        else
        {
            Console.WriteLine($"⚠ El campo '{campo}' no existe en ItemFormateado");
        }
    }

    return (itemFormateado, campos);
}

async Task ProcesarEnSolr(
    int itemCambiosId,
    string activo,
    ListItem item,
    JsonElement root,
    SolrManager solrManager,
    ClientContext context,
    LogsManager logsManager,
    bool incluirMilvus)
{
    try
    {
        // Obtener item formateado y campos modificados
        var (itemFormateado, campos) = await ObtenerItemFormateadoYCampos(item, root, solrManager, context);

        // Actualizar en Solr
        var resultadoSolr = await solrManager.AtomicUpdateSolrMultiple(
            item["GUID"].ToString(),
            campos,
            item.Id
        );

        if (resultadoSolr.ok)
        {
            Console.WriteLine($"✔️ Solr actualizado para ID {item.Id}");
            await logsManager.LogEjecucionAsync(itemCambiosId, $"Activo modificado en Solr (ID {item.Id})", "OK");

        }
        else
        {
            Console.WriteLine($"⚠️ Error Solr (ID {item.Id}): {resultadoSolr.msg}");
            await logsManager.LogErrorAsync(itemCambiosId, activo, $"Error al actualizar en Solr: {resultadoSolr.msg}", "Error Solr");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"⚠️ Excepción en Solr: {ex.Message}");
        await logsManager.LogErrorAsync(itemCambiosId, activo, $"Excepción al procesar en Solr: {ex.Message}", "Error Solr");
    }
}

async Task ProcesarEnMilvusAsync(
    HttpClient client,
    int itemCambiosId,
    string activo,
    int itemId,
    ListItem item,
    JsonElement root,
    SolrManager solrManager,
    ClientContext context,
    LogsManager logsManager,
    bool useIA)
{
    try
    {
        // Obtener item formateado y campos modificados (reutiliza la misma lógica que Solr)
        var (itemFormateado, campos) = await ObtenerItemFormateadoYCampos(item, root, solrManager, context);

        string json = JsonConvert.SerializeObject(itemFormateado, Newtonsoft.Json.Formatting.Indented);

        // Enviar a Milvus
        var milvusResult = await MilvusManager.EnviarAMilvus(client, json, itemId, useIA);

        if (milvusResult.ok)
        {
            Console.WriteLine($"✔️ Milvus actualizado para ID {itemId}");
            await logsManager.LogEjecucionAsync(itemCambiosId, $"Activo modificado en Milvus (ID {itemId})", "OK");
        }
        else
        {
            Console.WriteLine($"⚠️ Error al actualizar en Milvus: {milvusResult.msg}");
            await logsManager.LogErrorAsync(itemCambiosId, activo, $"Error al actualizar en Milvus: {milvusResult.msg}", "Error Milvus");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"⚠️ Excepción al actualizar en Milvus: {ex.Message}");
        await logsManager.LogErrorAsync(itemCambiosId, activo, $"Excepción al actualizar en Milvus: {ex.Message}", "Error Milvus");
    }
}