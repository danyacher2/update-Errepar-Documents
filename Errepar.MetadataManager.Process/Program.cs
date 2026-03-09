using Errepar.MetadataManager.Process.Auth;
using Errepar.MetadataManager.Process.Models;
using Errepar.MetadataManager.Process.Services;
using Errepar.MetadataManager.Process.Services;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client;
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


    // Autenticación con certificado usando thumbprint
    // Obtén el thumbprint desde el certificado instalado en Windows (ver pasos más abajo)
    string certificateThumbprint = Environment.GetEnvironmentVariable("CERT_THUMBPRINT")
        ?? "452079A2697BC9646023FAE02876488654BBDB2C"; // Reemplaza con tu thumbprint real

    // Generar token con certificado (thumbprint)
    var authToken = await TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync(tenantId, clientId, certificateThumbprint, siteUrl);
    using var http = new HttpClient();

    var context = new ClientContext(siteUrl);

    context.ExecutingWebRequest += (sender, e) =>
    {
        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authToken;
    };

    var logsManager = new LogsManager(http, siteUrl, "Metadata Manager");

    var searchService = new SearchService(context);

    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);

    var sp = new SharePointManager(http, siteUrl, "Metadata Manager", authToken);
    using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(2));

    var items = await sp.GetPendingOrPausedAsync(cts.Token).ConfigureAwait(false);

    var cambios = new ProcesarCambios(http, context, siteUrl, logsManager, searchService);

    context.Load(context.Web, w => w.ServerRelativeUrl);
    context.ExecuteQuery();

    var webServerRelative = context.Web.ServerRelativeUrl;
    Console.WriteLine("ServerRelativeUrl del sitio: " + webServerRelative);


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
    var activosEncontrados = new List<string>();

    Console.WriteLine($"Items obtenidos: {items.Count}");
    //foreach (var it in items)
    for (int j = 0; j < items.Count; j++)
            {
        var it = items[j];
        Console.WriteLine($"Id={it.Id} | Título='{it.Titulo}' | Estado='{it.EstadoProceso}' | EjecutadoPor='{it.EjecutadoPor}' | Link='{it.Link}'");
        //for  (var activo in it.Activos)
        for (int i = 0; i < it.Activos.Count; i++)
            {
            var activo = it.Activos[i];
            //var hit = searchService.ObtenerElementoPorGuid(activo);

            //if (hit != null)
            //{
            //Console.WriteLine("Path encontrado: " + hit.Path);

            //var siteUri = new Uri(siteUrl);
            //var hitUri = new Uri(hit.Path);

            //var serverRelativeUrl = new Uri(hit.Path).AbsolutePath;
            //.Replace("/sites/Errepar", "/sites/ErreparDesarrollo");

            // Validación defensiva
            //if (!serverRelativeUrl.StartsWith(siteUri.AbsolutePath, StringComparison.OrdinalIgnoreCase))
            //{
            //    Console.WriteLine("El item no pertenece a la misma url");
            //}
            //else
            //{



            //var file = context.Web.GetFileByServerRelativeUrl(serverRelativeUrl);
            //var file = context.Web.GetFileByServerRelativeUrl("/sites/ErreparDesarrollo/Documento/00-sin-obra/2024/20240126103422795/20240126103422795.html");


            //try
            //{

                Console.WriteLine($"{i} --Guid item: {activo}");

                var item = FindItem(listas, new Guid(activo), context, ref listaProbable);
                //var item = GetItemByUniqueId(context,new Guid(activo));


                //GetItemByUniqueId
                //Thread.Sleep(1000);

                //Console.WriteLine($"LIST: {listas} con el Guid {activo} y la lista probable {listaProbable}");
                Console.WriteLine($"item: {item}");
                //listas, Guid uniqueId, ClientContext ctx, List listaProbable)
                
                if (item != null)
                {
                    activosEncontrados.Add(activo);
                }
                

                continue;

                if (item == null)
                {
                    Console.WriteLine("❌ No se encontró el elemento en las listas candidatas");
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
                            cambios.Procesar(item, cambio);
                        }
                    }
                    else if (root.ValueKind == JsonValueKind.Object)
                    {
                        cambios.Procesar(item, root);
                    }
                    item.SystemUpdate();
                    context.ExecuteQuery();

                }
                var json = JsonSerializer.Serialize(item.FieldValues, jsonOptions);

                var baseFolder = Path.Combine(AppContext.BaseDirectory, "searchLogs");
                var listItemsFolder = Path.Combine(baseFolder, "ListItems");

                Directory.CreateDirectory(listItemsFolder); // crea toda la estructura

                var filePathNewItem = Path.Combine(listItemsFolder, $"Search_{activo}.json");

                System.IO.File.WriteAllText(filePathNewItem, json, Encoding.UTF8);

                Console.WriteLine($"✅ JSON completo guardado en: {filePathNewItem}");

                //}

                //}
                //else
                //{
                //    Console.WriteLine("No se encontró el elemento");
                //}

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
                    LibraryName = item.FieldValues["Title"].ToString(),
                    TimeStamp = Convert.ToInt32(item.FieldValues["ID"]),
                    UrlItem = item.FieldValues["FileRef"].ToString(),
                    Activo = activo,
                    Procesado = true,
                    Fecha = DateTime.Now
                };

                await logsManager.SaveLogActivosProcesados(it.Id, new[] { logActivo });
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("Error en la busqueda del Item"+ex);
            //}

            continue;

            //context.Load(file, f => f.ListItemAllFields, f => f.Exists);
            //context.ExecuteQuery();

            //if (!file.Exists)
            //{
            //    Console.WriteLine("❌ El path no corresponde a un archivo");
            //    return;
            //}

            //var item = file.ListItemAllFields;

            

        }


        //        await sp.UpdateListItemAsync(
        //    it.Id,
        //    new Dictionary<string, object>
        //    {
        //        { "EstadoProceso", "Finalizado" }
        //    }
        //);
        //await logsManager.SyncLogs(it.Id, cts.Token);


    }

    var folderActivos = Path.Combine(AppContext.BaseDirectory, "");
    Directory.CreateDirectory(folderActivos);

    var fileActivosEncontrados = Path.Combine(folderActivos, "ActivosEncontrados.txt");
    System.IO.File.WriteAllLines(fileActivosEncontrados, activosEncontrados, Encoding.UTF8);
    Console.WriteLine($"✅ Activos encontrados guardados en: {fileActivosEncontrados} | Cantidad: {activosEncontrados.Count}");
    Console.WriteLine("Presiona una tecla para iniciar...");
    Console.ReadKey();

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

static ListItem GetItemByUniqueId(ClientContext ctx,List listaProbable, Guid uniqueId)
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
                return items[0];
        }
        catch
        {
            // ignorar listas que fallen
        }
    

    return null;
}


ListItem FindItem(List<List> listas, Guid guid, ClientContext ctx, ref List listaProbable)
{
    //if (listaProbable != null)
    //{
    //    Console.WriteLine("LISTA: " + listaProbable.ToString());

    //    var item = GetItemByUniqueId(ctx, listaProbable, guid);
    //    if (item != null)
    //        return item;
    //}

    // 3️⃣ buscar en las demás listas
    foreach (var lista in listas)
    {
        //ctx.Load(lista, l => l.Id);
        //ctx.ExecuteQuery();
        //if (listaProbable != null && lista.Id == listaProbable.Id)
        //    continue;

        var item = GetItemByUniqueId(ctx, lista, guid);       

        if (item != null)
        {
            listaProbable = lista; // guardar para próximas búsquedas
            return item;
        }
    }

    return null;
}

ListItem BuscarEnLista(List lista, Guid guid, ClientContext ctx)
{
    var query = new CamlQuery
    {
        ViewXml = $@"
    <View Scope='RecursiveAll'>
      <ViewFields>
        <FieldRef Name='ID' />
      </ViewFields>
      <Query>
        <Where>
          <Eq>
            <FieldRef Name='GUID'/>
            <Value Type='Guid'>{guid}</Value>
          </Eq>
        </Where>
      </Query>
      <RowLimit>10</RowLimit>
    </View>"
    };

    var items = lista.GetItems(query);

    ctx.Load(items, col => col.Include(i => i.Id));
    ctx.ExecuteQuery();

    return items.Count > 0 ? items[0] : null;
}



//ListItem FindItem(List<List> listas, Guid uniqueId, ClientContext ctx, ref List listaProbable)
//{
//    var item = BuscarEnLista(listaProbable, uniqueId, ctx);
//    if (item != null)
//    {
//        return item;

//    }

//    foreach (var lista in listas)
//    {
//        if (lista.Id == listaProbable.Id)
//            continue;

//        item = BuscarEnLista(lista, uniqueId, ctx);
//        if (item != null)
//        {
//            listaProbable = lista;
//            return item;

//        }
//    }

//    return null;
//}

//ListItem BuscarEnLista(List lista, Guid uniqueId, ClientContext ctx)
//{
//    var query = new CamlQuery
//    {
//        ViewXml = $@"
//    <View Scope='RecursiveAll'>
//        <Query>
//            <Where>
//                <Eq>
//                    <FieldRef Name='id'/>
//                    <Value Type='Guid'>{uniqueId}</Value>
//                </Eq>
//            </Where>
//        </Query>
//        <RowLimit>1</RowLimit>
//    </View>"
//    };

//    var items = lista.GetItems(query);
//    ctx.Load(items);
//    ctx.ExecuteQuery();

//    return items.Count > 0 ? items[0] : null;
//}



static string Base64Encode(string plainText)
{
    var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
    return System.Convert.ToBase64String(plainTextBytes);
}

static SecureString FetchPasswordFromConsole(string pass)
{
    string password = pass;
    var securePassword = new SecureString();
    //Convert string to secure string  
    foreach (char c in password)
        securePassword.AppendChar(c);
    securePassword.MakeReadOnly();
    return securePassword;
}


/*
// Helpers
static string Prompt(string message)
{
    Console.Write(message);
    return Console.ReadLine() ?? string.Empty;
}

static string PromptSecret(string message)
{
    Console.Write(message);
    var sb = new StringBuilder();
    ConsoleKeyInfo key;
    while ((key = Console.ReadKey(true)).Key != ConsoleKey.Enter)
    {
        if (key.Key == ConsoleKey.Backspace && sb.Length > 0)
        {
            sb.Length--;
            Console.Write("\b \b");
        }
        else if (!char.IsControl(key.KeyChar))
        {
            sb.Append(key.KeyChar);
            Console.Write('*');
        }
    }
    Console.WriteLine();
    return sb.ToString();
}*/