using Errepar.MetadataManager.Process.Auth;
using Errepar.MetadataManager.Process.Config;
using Errepar.MetadataManager.Process.Models;
using Errepar.MetadataManager.Process.Services;
using ClosedXML.Excel;
using Microsoft.SharePoint.Client;
using System.Text.Encodings.Web;
using System.Text.Json;

if (args.Any(argument => string.Equals(argument, "--convert-xlsx", StringComparison.OrdinalIgnoreCase)))
{
    var xlsxPath = GetOptionValue(args, "--xlsx") ?? FindFile("resultado.xlsx");
    if (xlsxPath == null || !System.IO.File.Exists(xlsxPath))
        throw new FileNotFoundException("No se encontró el XLSX. Usá --xlsx=RUTA para indicar el archivo.");

    xlsxPath = Path.GetFullPath(xlsxPath);
    var sheetName = GetOptionValue(args, "--sheet");
    var filterColumn = GetOptionValue(args, "--filter-column");
    var filterValue = GetOptionValue(args, "--filter-value");
    if ((filterColumn == null) != (filterValue == null))
        throw new ArgumentException("Usá juntos --filter-column y --filter-value.");

    var defaultJsonPath = sheetName == null
        ? Path.ChangeExtension(xlsxPath, ".json")
        : Path.Combine(
            Path.GetDirectoryName(xlsxPath)!,
            $"{Path.GetFileNameWithoutExtension(xlsxPath)}-{MakeSafeFileName(sheetName)}{(filterValue == null ? string.Empty : $"-{MakeSafeFileName(filterValue)}")}.json");
    var jsonPath = GetOptionValue(args, "--json") ?? defaultJsonPath;
    var json = GenerateJsonFromExcel(xlsxPath, sheetName, filterColumn, filterValue);
    await System.IO.File.WriteAllTextAsync(jsonPath, json);
    Console.WriteLine($"JSON generado: {Path.GetFullPath(jsonPath)}");
    return;
}

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


static string? GetOptionValue(string[] arguments, string optionName)
{
    var prefix = optionName + "=";
    var argument = arguments.FirstOrDefault(value => value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
    return argument == null ? null : argument[prefix.Length..];
}

static string GenerateJsonFromExcel(
    string xlsxPath,
    string? sheetName = null,
    string? filterColumn = null,
    string? filterValue = null)
{
    using var xlsxStream = new FileStream(
        xlsxPath,
        FileMode.Open,
        FileAccess.Read,
        FileShare.ReadWrite | FileShare.Delete);
    using var workbook = new XLWorkbook(xlsxStream);
    var worksheet = sheetName == null
        ? workbook.Worksheets.FirstOrDefault()
        : workbook.Worksheets.FirstOrDefault(
            sheet => string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase));
    if (worksheet == null)
        throw new InvalidDataException(
            sheetName == null
                ? "El archivo XLSX no contiene hojas."
                : $"No se encontró la hoja '{sheetName}'. Hojas disponibles: {string.Join(", ", workbook.Worksheets.Select(sheet => sheet.Name))}.");

    var firstRow = worksheet.FirstRowUsed()?.RowNumber()
        ?? throw new InvalidDataException("La primera hoja del XLSX está vacía.");
    var lastRow = worksheet.LastRowUsed()!.RowNumber();
    var firstColumn = worksheet.FirstColumnUsed()!.ColumnNumber();
    var lastColumn = worksheet.LastColumnUsed()!.ColumnNumber();

    var headers = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1)
        .Select(column => worksheet.Cell(firstRow, column).GetString().Trim())
        .ToArray();

    var filterColumnIndex = -1;
    if (filterColumn != null)
    {
        filterColumnIndex = Array.FindIndex(
            headers,
            header => string.Equals(header, filterColumn, StringComparison.OrdinalIgnoreCase));
        if (filterColumnIndex < 0)
            throw new InvalidDataException(
                $"No se encontró la columna '{filterColumn}' en la hoja '{worksheet.Name}'. Encabezados: {string.Join(" | ", headers)}.");
    }

    var emptyHeaderIndex = Array.FindIndex(headers, string.IsNullOrWhiteSpace);
    if (emptyHeaderIndex >= 0)
        throw new InvalidDataException($"El encabezado de la columna {firstColumn + emptyHeaderIndex} está vacío.");

    var duplicateHeader = headers
        .GroupBy(header => header, StringComparer.OrdinalIgnoreCase)
        .FirstOrDefault(group => group.Count() > 1);
    if (duplicateHeader != null)
        throw new InvalidDataException($"El encabezado '{duplicateHeader.Key}' está repetido.");

    var records = new List<Dictionary<string, object?>>();
    for (var row = firstRow + 1; row <= lastRow; row++)
    {
        var cells = Enumerable.Range(firstColumn, headers.Length)
            .Select(column => worksheet.Cell(row, column))
            .ToArray();
        if (cells.All(cell => cell.IsEmpty()))
            continue;

        if (filterColumnIndex >= 0
            && !cells[filterColumnIndex].GetString().Trim().StartsWith(filterValue!, StringComparison.OrdinalIgnoreCase))
            continue;

        var record = new Dictionary<string, object?>(StringComparer.Ordinal);
        for (var index = 0; index < headers.Length; index++)
            record[headers[index]] = GetJsonCellValue(cells[index]);

        records.Add(record);
    }

    return JsonSerializer.Serialize(records, new JsonSerializerOptions
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    });
}

static string MakeSafeFileName(string value)
{
    var invalidCharacters = Path.GetInvalidFileNameChars();
    return string.Join("_", value.Split(invalidCharacters, StringSplitOptions.RemoveEmptyEntries));
}

static object? GetJsonCellValue(IXLCell cell)
{
    if (cell.IsEmpty())
        return null;

    return cell.DataType switch
    {
        XLDataType.Boolean => cell.GetBoolean(),
        XLDataType.Number => cell.GetDouble(),
        XLDataType.DateTime => cell.GetDateTime(),
        XLDataType.TimeSpan => cell.GetTimeSpan(),
        XLDataType.Error => cell.GetError().ToString(),
        _ => cell.GetString()
    };
}

static string? FindDataFile() => FindFile("data.json");

static string? FindFile(string fileName)
{
    foreach (var startDirectory in new[] { Directory.GetCurrentDirectory(), AppContext.BaseDirectory })
    {
        for (var directory = new DirectoryInfo(startDirectory); directory != null; directory = directory.Parent)
        {
            var candidate = Path.Combine(directory.FullName, fileName);
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