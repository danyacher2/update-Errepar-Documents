using Errepar.MetadataManager.Process.Auth;
using Errepar.MetadataManager.Process.Models;
using Errepar.MetadataManager.Process.Services;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.News.DataModel;
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

// Configuración global para resolver problemas de DNS con VPN
ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
ServicePointManager.DnsRefreshTimeout = 0; // Forzar resolución DNS en cada intento
ServicePointManager.EnableDnsRoundRobin = true;

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
    string clientId = "f679c472-7b0c-45dc-b38c-cca0b662f77a";


    string certificateThumbprint = Environment.GetEnvironmentVariable("CERT_THUMBPRINT")
        ?? "452079A2697BC9646023FAE02876488654BBDB2C";
    var authToken = await TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync(tenantId, clientId, certificateThumbprint, siteUrl);
    using var http = new HttpClient();
    http.Timeout = TimeSpan.FromMinutes(10);

    var context = new ClientContext(siteUrl);

    var solerManager = new SolrManager(http, siteUrl, "UAT");


    context.ExecutingWebRequest += (sender, e) =>
    {
        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authToken;
    };

    var logsManager = new LogsManager(http, siteUrl, "Metadata Manager");


    await logsManager.SaveLogEjecucion(0, new ItemLogEjecucion
    {
        ItemId = 0,
        Fecha = DateTime.Now,
        Mensaje = "Inicio del proceso Metadata Manager",
        Estado = "Inicio"
    });

    var searchService = new SearchService(context);

    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);

    var sp = new SharePointManager(http, siteUrl, "Metadata Manager", authToken);
    using var cts = new CancellationTokenSource(TimeSpan.FromHours(10));

    var items = await sp.GetPendingOrPausedAsync(cts.Token).ConfigureAwait(false);
    await logsManager.SaveLogEjecucion(0, new ItemLogEjecucion
    {
        ItemId = 0,
        Fecha = DateTime.Now,
        Mensaje = $"Items recuperados desde SharePoint: {items.Count}",
        Estado = "OK"
    });



    var cambios = new ProcesarCambios(http, context, siteUrl, logsManager, searchService);

    context.Load(context.Web, w => w.ServerRelativeUrl);
    context.ExecuteQuery();

    var webServerRelative = context.Web.ServerRelativeUrl;
    Console.WriteLine("ServerRelativeUrl del sitio: " + webServerRelative);

    var (allow, deny, rulesCount) = ScheduleGate.LoadWindows(context, "Manage Metadata Schedule");

    await logsManager.SaveLogEjecucion(0, new ItemLogEjecucion
    {
        ItemId = 0,
        Fecha = DateTime.Now,
        Mensaje = $"Reglas de Schedule cargadas: {rulesCount}",
        Estado = "OK"
    });

    var EstadoProceso = "Procesando";

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
        var listName = context.Web.Lists.GetByTitle(name);
        context.Load(listName);
        listas.Add(listName);
    }
    var listaProbable = listas.FirstOrDefault();
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
            //try
            //{
            if (!ScheduleGate.IsAllowed(DateTimeOffset.Now, allow, deny))
            {
                if (EstadoProceso != "En Pausa")
                {
                    EstadoProceso = "En Pausa";

                    await sp.UpdateListItemAsync(
                        it.Id,
                        new Dictionary<string, object>
                        {
                { "EstadoProceso", "En Pausa" }
                        });
                }
                Console.WriteLine("Proceso en Pausa");

                await logsManager.SaveLogEjecucion(it.Id, new ItemLogEjecucion
                {
                    ItemId = it.Id,
                    Fecha = DateTime.Now,
                    Mensaje = "Proceso en Pausa",
                    Estado = "En Pausa"
                });

                await ScheduleGate.WaitUntilAllowedAsync(allow, deny, TimeSpan.FromMinutes(1), cts.Token);

                EstadoProceso = "Procesando";

                await sp.UpdateListItemAsync(
                    it.Id,
                    new Dictionary<string, object>
                    {
            { "EstadoProceso", "Procesando" }
                                 });

                await logsManager.SaveLogEjecucion(it.Id, new ItemLogEjecucion
                {
                    ItemId = it.Id,
                    Fecha = DateTime.Now,
                    Mensaje = "El Procesndo",
                    Estado = "Procesando"
                });

                Console.WriteLine("Proceso Procesando");

            }




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
                await logsManager.SaveLogError(it.Id, new ItemLogError
                {
                    ItemId = activo,
                    Fecha = DateTime.Now,
                    Mensaje = $"Item con guid {activo} no se encontro",
                    Estado = "No Encontrado"
                });
                continue;
            }

            var jsonOldItem = JsonSerializer.Serialize(item.FieldValues, jsonOptions);

            var folder = Path.Combine(AppContext.BaseDirectory, "searchLogs");
            Directory.CreateDirectory(folder);

            var filePath = Path.Combine(folder, $"ListItems/SearchOriginal_{activo}.json");
            System.IO.File.WriteAllText(filePath, jsonOldItem, Encoding.UTF8);
            Console.WriteLine($"✅ JSON completo guardado en: {filePath}");

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


                //Soler
                //tengo que armar 
                var campos = new Dictionary<string, object>();
                //{
                //    { "eolShpTimeStamp", timestamp },
                //    { "eolShpFechaBusqueda", fechaFormateada }
                //};

                try
                {
                    IDictionary<string, object> ItemFormateado = await solerManager.GetItemFormatSync(context, item);

                    foreach (JsonElement cambio in root.EnumerateArray())
                    {
                        var campo = GetString(cambio, "campo");

                        if (ItemFormateado.TryGetValue(campo, out var valor))
                        {
                            if (!campos.ContainsKey(campo))
                                campos.Add(campo, valor);
                        }
                        else
                        {
                            Console.WriteLine($"⚠ El campo '{campo}' no existe en ItemFormateado");
                        }
                    }

                    var resultadoSolr = await solerManager.AtomicUpdateSolrMultiple(item["GUID"].ToString(), campos, item.Id);

                    //if (resultadoSolr.ok)
                    //{
                    //    Console.WriteLine($"✔️ Solr actualizado para ID {item.Id}");
                    //    await logsManager.SaveLogError(it.Id, new ItemLogError
                    //    {
                    //        ItemId = activo,
                    //        Fecha = DateTime.Now,
                    //        Mensaje = resultadoSolr.msg,
                    //        Estado = "Activo modificado en Solr"
                    //    });
                    //}

                    //else
                    //{
                    //    Console.WriteLine($"⚠️ Error Solr (ID {item.Id}): {resultadoSolr.msg}");
                    //    await logsManager.SaveLogError(it.Id, new ItemLogError
                    //    {
                    //        ItemId = activo,
                    //        Fecha = DateTime.Now,
                    //        Mensaje = resultadoSolr.msg,
                    //        Estado = "Activo no encontrado en Solr"
                    //    });
                    //    //Guardar en log que no se encontro 
                    //}

                    try
                    {
                        //Milvus 
                        //string json = JsonConvert.SerializeObject(col, Newtonsoft.Json.Formatting.Indented);
                        //
                    }
                    catch (Exception ex)
                    {
                        //no se pudo grabar en milvus
                    }

                }
                catch (Exception ex)
                {
                    //no se pudo grabar en soler
                }


            }
            CantidadProcesdos++;
            var json = JsonSerializer.Serialize(item.FieldValues, jsonOptions);

            var baseFolder = Path.Combine(AppContext.BaseDirectory, "searchLogs");
            var listItemsFolder = Path.Combine(baseFolder, "ListItems");

            Directory.CreateDirectory(listItemsFolder);

            var filePathNewItem = Path.Combine(listItemsFolder, $"Search_{activo}.json");

            System.IO.File.WriteAllText(filePathNewItem, json, Encoding.UTF8);

            Console.WriteLine($"✅ JSON completo guardado en: {filePathNewItem}");

            await logsManager.SaveLogEjecucion(it.Id, new ItemLogEjecucion
            {
                ItemId = it.Id,
                Fecha = DateTime.Now,
                Mensaje = "Inicio proceso",
                Estado = "OK"
            });



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

            await logsManager.SaveLogError(it.Id, new ItemLogError
            {
                ItemId = activo,
                Fecha = DateTime.Now,
                Mensaje = $"Item con guid {activo} no Modificado",
                Estado = "error en el cambio"
            });
            //    Console.WriteLine("Error en la busqueda del Item"+ex);
            //}
            await sp.UpdateListItemAsync(
    it.Id,
    new Dictionary<string, object>
    {
{ "CantActivosProcesados", CantidadProcesdos }
                 });
        }


        await sp.UpdateListItemAsync(
                it.Id,
                new Dictionary<string, object>
                {
                    { "EstadoProceso", "Finalizado"
                    }
                }
                );


        await logsManager.SyncLogs(it.Id, cts.Token);

    }

    await logsManager.SaveLogEjecucion(0, new ItemLogEjecucion
    {
        ItemId = 0,
        Fecha = DateTime.Now,
        Mensaje = "Proceso completado correctamente",
        Estado = "Finalizado"
    });

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
            ctx.Load(item, i => i.ContentType, i => i.ParentList, i => i.ParentList.Fields, i => i["eolShpTema"], i => i["eolShpCodigoMislibros"]);
            ctx.ExecuteQuery();
            return item;
        }
    }
    catch
    {
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