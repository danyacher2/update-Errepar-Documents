using Errepar.MetadataManager.Process.Auth;
using Errepar.MetadataManager.Process.Models;
using Errepar.MetadataManager.Process.Services;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.News.DataModel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Linq;
using System.Net;
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
string siteUrl = "https://erreparsa.sharepoint.com/sites/ErreparDesarrollo";
string tenantId = "00f26ad1-2073-4746-a79f-c83061db35c0";
string clientId = "f679c472-7b0c-45dc-b38c-cca0b662f77a";//erreparDev    //string clientId = "8688eed4-7464-4288-9820-34849fd19296";//erreparDesarrollo

string certificateThumbprint = Environment.GetEnvironmentVariable("CERT_THUMBPRINT")
    ?? "452079A2697BC9646023FAE02876488654BBDB2C";
var authToken = await TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync(tenantId, clientId, certificateThumbprint, siteUrl);
DateTime tokenTime = DateTime.MinValue;

// Delegate para actualizar el token en ClientContext
void UpdateContextToken(object sender, WebRequestEventArgs e)
{
    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authToken;
}

ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
ServicePointManager.DnsRefreshTimeout = 0; // Forzar resolución DNS en cada intento
ServicePointManager.EnableDnsRoundRobin = true;


try
{
    Console.WriteLine("Consulta a SharePoint: recuperar items con EstadoProceso = Pendiente | En Pausa");


    using var http = new HttpClient();
    http.Timeout = TimeSpan.FromMinutes(10);

    var context = new ClientContext(siteUrl);

    var solerManager = new SolrManager(http, siteUrl, "UAT");

    context.ExecutingWebRequest += UpdateContextToken;
    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);

    var logsManager = new LogsManager(http, siteUrl, "Metadata Manager");

    await logsManager.LogEjecucionAsync(0, "Inicio del proceso Metadata Manager", "Inicio");

    var searchService = new SearchService(context);

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
        await sp.UpdateEstadoProcesoAsync(it.Id, "Procesando");
        if (it.CantActivosProcesados != null && it.CantActivosProcesados+1 < it.CantActivosSeleccionados && it.CantActivosProcesados > 0)
        {
            CantidadProcesdos = it.CantActivosProcesados.Value;
            if (CantidadProcesdos > 0)
                CantidadProcesdos--;
        }
        Console.WriteLine($"Id={it.Id} | Título='{it.Titulo}' | Estado='{it.EstadoProceso}' | EjecutadoPor='{it.EjecutadoPor}' | Link='{it.Link}'");
        int i= 0;
        for (i = CantidadProcesdos; i < it.Activos.Count; i++)
        {
            try
            {
                EstadoProceso = await ManejarPausaPorSchedule(EstadoProceso, it.Id, allow, deny, sp, logsManager, cts.Token);

                // Renovar token para cada activo (asegurar que esté válido)
                authToken = await GetValidTokenAsync();
                http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);

                // El delegate UpdateContextToken usa la variable authToken actualizada automáticamente

                var activo = it.Activos[i];

                var procesarEnShpSolrMilvus = it.Scope;

                Console.WriteLine($"{i} --Guid item: {activo}");


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

                    try
                    {
                        var (itemFormateado, campos) = await ObtenerItemFormateadoYCampos(item, root, solerManager, context);

                        if (procesarEnShpSolrMilvus.Contains("Soler"))
                        {
                            bool incluirMilvus = procesarEnShpSolrMilvus.Contains("Milvus");
                            await ProcesarEnSolr(it.Id, activo, item, campos, solerManager, context, logsManager, incluirMilvus, sp);
                        }

                        if (procesarEnShpSolrMilvus.Contains("Milvus - IA"))
                        {
                            await ProcesarEnMilvusAsync(http, it.Id, activo, item.Id, item, itemFormateado, solerManager, context, logsManager, true, sp);
                        }
                        else if (procesarEnShpSolrMilvus.Contains("Milvus"))
                        {
                            await ProcesarEnMilvusAsync(http, it.Id, activo, item.Id, item, itemFormateado, solerManager, context, logsManager, false, sp);
                        }
                    }
                    catch (Exception ex) when (ex.Message.Contains("Error crítico en Solr") || ex.Message.Contains("Error crítico en Milvus"))
                    {
                        // Re-lanzar excepciones críticas para detener el proceso
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠️ Excepción al actualizar en SharePoint: {ex}");
                        await logsManager.LogErrorAsync(it.Id, activo, $"Excepción al actualizar en SharePoint: {ex}", "Error SharePoint");
                    }
                    await logsManager.SaveItemProcessedJsonAsync(activo, item.FieldValues);

                    await logsManager.LogEjecucionAsync(it.Id, "Item actualizado correctamente:" + item.Id, "OK");

                    var logActivo = new ItemLogActivosProcesados
                    {
                        ItemCambiosId = it.Id,
                        ItemListId = Convert.ToInt32(item.FieldValues["ID"]),
                        LibraryName = item.FieldValues["Title"]?.ToString(),
                        TimeStamp = item.FieldValues["eolShpTimeStamp"].ToString(),
                        UrlItem = item.FieldValues["FileRef"]?.ToString(),
                        Activo = activo,
                        Procesado = true,
                        Fecha = DateTime.Now
                    };

                    await logsManager.SaveLogActivosProcesados(it.Id, new[] { logActivo });
                }
                try
                {
                    // Renovar token antes de actualizar SharePoint
                    authToken = await GetValidTokenAsync();
                    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);

                    // Actualizar también el contexto CSOM
                    context.ExecutingWebRequest -= UpdateContextToken;
                    context.ExecutingWebRequest += UpdateContextToken;

                    await sp.UpdateCantActivosAsync(it.Id, i++);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error actualizando cantidad de activos: " + ex);
                }
            }
            catch (Exception ex) when (ex.Message.Contains("Error crítico en Solr") || ex.Message.Contains("Error crítico en Milvus"))
            {
                // Re-lanzar excepciones críticas para detener el proceso
                Console.Error.WriteLine($"❌ Excepción crítica detectada: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                await logsManager.LogErrorAsync(it.Id, it.Id.ToString(), $"Item con guid {it.Id} no Modificado", "error en el cambio" + ex);
                Console.WriteLine("Error en la busqueda del Item" + ex);
            }
            
        }
        await sp.UpdateCantActivosAsync(it.Id, i++);
        try
        {
            // Renovar token antes de finalizar el item
            authToken = await GetValidTokenAsync();
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);

            context.ExecutingWebRequest -= UpdateContextToken;
            context.ExecutingWebRequest += UpdateContextToken;

            await sp.UpdateEstadoProcesoAsync(it.Id, "Finalizado");
            await logsManager.SyncLogs(it.Id, cts.Token);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error finalizando item: " + ex);
                    }
    }
    try
    {
        await logsManager.LogEjecucionAsync(0, "Proceso completado correctamente", "Finalizado");

    }
    catch (Exception ex)
    {
        Console.WriteLine("Error" + ex);
    }
}

catch (Exception ex)
{
    Console.Error.WriteLine("Excepción no controlada: " + ex);
    //Estado error
    Environment.ExitCode = 1;
}
finally
{
    Console.WriteLine("Proceso finalizado con código " + Environment.ExitCode + ". Presiona una tecla para cerrar...");
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
    catch (Exception Ex)
    {
        Console.WriteLine("Error" + Ex);
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
    int itemId,
    string activo,
    ListItem item,
    Dictionary<string, object> campos,
    SolrManager solrManager,
    ClientContext context,
    LogsManager logsManager,
    bool incluirMilvus,
    SharePointManager sp)
{
    try
    {
        // Actualizar en Solr
        var resultadoSolr = await solrManager.AtomicUpdateSolrMultiple(
            item["GUID"].ToString(),
            campos,
            item.Id
        );

        if (resultadoSolr.ok)
        {
            Console.WriteLine($"✔️ Solr actualizado para ID {item.Id}");
            await logsManager.LogEjecucionAsync(itemId, $"Activo modificado en Solr (ID {item.Id})", "OK");
        }
        else
        {
            // ❌ Falló después de 3 intentos
            Console.WriteLine($"❌ ERROR CRÍTICO Solr (ID {item.Id}): {resultadoSolr}");
            await logsManager.LogErrorAsync(itemId, activo, $"Error crítico en Solr después de 3 intentos: {resultadoSolr}", "Error Solr");

            // Actualizar estado en SharePoint
            await sp.UpdateEstadoProcesoAsync(itemId, "Error Solr");
            await logsManager.SyncLogs(itemId, default);

            Console.Error.WriteLine($"⛔ Proceso detenido por error en Solr para item {itemId}");
            Environment.ExitCode = 2; // Código específico para error Solr
            throw new Exception($"Error crítico en Solr: {resultadoSolr}");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"⚠️ Excepción inesperada en Solr: {ex}");
        await logsManager.LogErrorAsync(itemId, activo, $"Excepción al procesar en Solr: {ex}", "Error Solr");
        await sp.UpdateEstadoProcesoAsync(itemId, "Error Solr");
        await logsManager.SyncLogs(itemId, default);
        Environment.ExitCode = 2;
        throw new Exception($"Excepción crítica en Solr: {ex}", ex);
    }
}

async Task ProcesarEnMilvusAsync(
    HttpClient client,
    int itemCambiosId,
    string activo,
    int itemId,
    ListItem item,
    IDictionary<string, object> itemFormateado,
    SolrManager solrManager,
    ClientContext context,
    LogsManager logsManager,
    bool useIA,
    SharePointManager sp)
{
    try
    {
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
            // ❌ Falló después de 3 intentos
            Console.WriteLine($"❌ ERROR CRÍTICO Milvus (ID {itemId}): {milvusResult}");
            await logsManager.LogErrorAsync(itemCambiosId, activo, $"Error crítico en Milvus después de 3 intentos: {milvusResult}", "Error Milvus");

            // Actualizar estado en SharePoint
            await sp.UpdateEstadoProcesoAsync(itemCambiosId, "Error Milvus");
            await logsManager.SyncLogs(itemCambiosId, default);

            Console.Error.WriteLine($"⛔ Proceso detenido por error en Milvus para item {itemCambiosId}");
            Environment.ExitCode = 3; // Código específico para error Milvus
            throw new Exception($"Error crítico en Milvus: {milvusResult}");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"⚠️ Excepción inesperada en Milvus: {ex}");
        await logsManager.LogErrorAsync(itemCambiosId, activo, $"Excepción al actualizar en Milvus: {ex}", "Error Milvus");
        await sp.UpdateEstadoProcesoAsync(itemCambiosId, "Error Milvus");
        await logsManager.SyncLogs(itemCambiosId, default);

        Environment.ExitCode = 3;
        throw new Exception($"Excepción crítica en Milvus: {ex}", ex);
    }
}



async Task<string> GetValidTokenAsync()
{
    // si no existe o pasaron 50 minutos → renovar
    if (authToken == null || (DateTime.UtcNow - tokenTime).TotalMinutes >= 50)
    {
        Console.WriteLine("🔄 Renovando token...");

        authToken = await TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync(
            tenantId, clientId, certificateThumbprint, siteUrl);

        tokenTime = DateTime.UtcNow;
    }

    return authToken;
}

