using Errepar.MetadataManager.Process.Auth;
using Errepar.MetadataManager.Process.Config;
using Errepar.MetadataManager.Process.Models;
using Errepar.MetadataManager.Process.Services;
using Microsoft.SharePoint.Client;
using System.Text.Json;

var environment = EnvironmentConfig.ResolveFromArgs(args);
var sharePoint = environment.SharePoint;
var authToken = await TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync(
    sharePoint.TenantId,
    sharePoint.ClientId,
    sharePoint.CertificateThumbprint,
    sharePoint.SiteUrl);

using var context = new ClientContext(sharePoint.SiteUrl);
context.ExecutingWebRequest += (_, request) =>
    request.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authToken;

var dataPath = FindDataFile();
if (dataPath == null)
    throw new FileNotFoundException("No se encontró data.json en el directorio actual ni en sus directorios superiores.");

using var document = JsonDocument.Parse(await System.IO.File.ReadAllTextAsync(dataPath));
if (document.RootElement.ValueKind != JsonValueKind.Array)
    throw new InvalidDataException("data.json debe contener un array de objetos con la propiedad 'id'.");

var lists = LoadSharePointLibraries(context);
context.ExecuteQuery();
var http = new HttpClient();
var solrManager = new SolrManager(http, environment.Solr, environment.Milvus);
var procesarEnShpSolrMilvus = new[] { "Solr" };
var cambios = new ProcesarCambios(
    http,
    context,
    sharePoint.SiteUrl,
    new LogsManager(http, sharePoint.SiteUrl, "update-errepar-document"),
    new SearchService(context));


foreach (var entry in document.RootElement.EnumerateArray())
{
    if (!entry.TryGetProperty("id", out var idElement) || !Guid.TryParse(idElement.GetString(), out var guid))
    {
        Console.WriteLine($"ID inválido en data.json: {entry}");
        continue;
    }

    try
    {
        var item = FindItem(lists, guid, context);
        if (item == null)
        {
            Console.WriteLine($"GUID {guid}: no encontrado");
            continue;
        }
        Console.WriteLine($"GUID {guid} | Título: {item["Title"]}");
        // var actualizado = cambios.AgregarIndice(item, "4e4b1d54-ac85-473d-a88b-8ccd867af570", "Compraventa");
        // Console.WriteLine(actualizado
        //     ? $"GUID {guid}: índice guardado en SharePoint"
        //     : $"GUID {guid}: sin cambios (índice existente o GUID de término vacío)");

        if (procesarEnShpSolrMilvus.Contains("Solr", StringComparer.OrdinalIgnoreCase))
        {
            context.Load(
                item,
                currentItem => currentItem.ContentType,
                currentItem => currentItem.ParentList,
                currentItem => currentItem.ParentList.Fields);
            context.ExecuteQuery();

            var itemFormateado = await solrManager.GetItemFormatSync(context, item);
            var payloadSolr = Newtonsoft.Json.JsonConvert.SerializeObject(new[] { itemFormateado });
            var resultadoSolr = await solrManager.EnviarASolr(payloadSolr, item.Id);

            Console.WriteLine(resultadoSolr.ok
                ? $"GUID {guid}: documento actualizado en Solr"
                : $"GUID {guid}: error al actualizar Solr: {resultadoSolr.msg}");

            // bool incluirMilvus = procesarEnShpSolrMilvus.Contains("Milvus", StringComparer.OrdinalIgnoreCase);
            // if (incluirMilvus)
            // {
            //     var payloadMilvus = Newtonsoft.Json.JsonConvert.SerializeObject(itemFormateado);
            //     var resultadoMilvus = await MilvusManager.EnviarAMilvus(
            //         http,
            //         payloadMilvus,
            //         item.Id,
            //         environment.Milvus,
            //         false);

            //     Console.WriteLine(resultadoMilvus.ok
            //         ? $"GUID {guid}: documento actualizado en Milvus"
            //         : $"GUID {guid}: error al actualizar Milvus: {resultadoMilvus.msg}");
            // }
        }
    }
    catch (Exception exception)
    {
        Console.WriteLine($"GUID {guid}: error durante la búsqueda: {exception.Message}");
    }
}


string? FindDataFile()
{
    foreach (var startDirectory in new[] { Directory.GetCurrentDirectory(), AppContext.BaseDirectory })
    {
        for (var directory = new DirectoryInfo(startDirectory); directory != null; directory = directory.Parent)
        {
            var candidate = Path.Combine(directory.FullName, "data.json");
            if (System.IO.File.Exists(candidate))
                return candidate;
        }
    }

    return null;
}

List<Microsoft.SharePoint.Client.List> LoadSharePointLibraries(ClientContext clientContext)
{
    var libraryNames = new[]
    {
        "Jurisprudencia Adm",
        "Documento",
        "Agenda",
        "Doctrina",
        "Guía Temática",
        "Jurisprudencia Judicial",
        "Legislación",
        "Modelo",
        "Actualidad",
        "Videos",
        "Podcast"
    };

    var lists = new List<Microsoft.SharePoint.Client.List>();
    foreach (var libraryName in libraryNames)
    {
        var list = clientContext.Web.Lists.GetByTitle(libraryName);
        clientContext.Load(list, currentList => currentList.Title, currentList => currentList.Id);
        lists.Add(list);
    }

    return lists;
}

static ListItem? FindItem(
    IEnumerable<Microsoft.SharePoint.Client.List> lists,
    Guid guid,
    ClientContext context)
{
    foreach (var list in lists)
    {
        var item = GetItemByUniqueIdPagedScan(context, list, guid);
        if (item != null)
            return item;
    }

    return null;
}

static ListItem? GetItemByUniqueIdPagedScan(ClientContext context, Microsoft.SharePoint.Client.List list, Guid uniqueId)
{
    var targetGuid = uniqueId.ToString();
    ListItemCollectionPosition? position = null;

    do
    {
        var query = new CamlQuery
        {
            ListItemCollectionPosition = position,
            ViewXml = @"
                <View Scope='RecursiveAll'>
                    <Query>
                        <OrderBy><FieldRef Name='ID' /></OrderBy>
                    </Query>
                    <ViewFields>
                        <FieldRef Name='ID' /><FieldRef Name='GUID' />
                    </ViewFields>
                    <RowLimit>2000</RowLimit>
                </View>"
        };

        var items = list.GetItems(query);
        context.Load(items, collection => collection.Include(item => item.Id, item => item["GUID"]), collection => collection.ListItemCollectionPosition);
        context.ExecuteQuery();

        foreach (var candidate in items)
        {
            var itemGuid = candidate["GUID"]?.ToString()?.Trim('{', '}');
            if (!string.Equals(itemGuid, targetGuid, StringComparison.OrdinalIgnoreCase))
                continue;

            context.Load(candidate);
            context.ExecuteQuery();
            return candidate;
        }

        position = items.ListItemCollectionPosition;
    }
    while (position != null);

    return null;
}