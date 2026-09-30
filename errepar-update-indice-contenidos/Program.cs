using Errepar.MetadataManager.Process.Auth;
using Errepar.MetadataManager.Process.Config;
using Errepar.MetadataManager.Process.Models;
using Errepar.MetadataManager.Process.Services;
using ClosedXML.Excel;
using Microsoft.SharePoint.Client;
using System.Text.Encodings.Web;
using System.Text.Json;

// Modo independiente: convierte un archivo Excel a JSON y termina sin conectarse a SharePoint.
if (args.Any(argument => string.Equals(argument, "--convert-xlsx", StringComparison.OrdinalIgnoreCase)))
{
    // Si no se especifica --xlsx=..., busca el archivo predeterminado en el directorio actual o sus superiores.
    var xlsxPath = GetOptionValue(args, "--xlsx") ?? FindFile("resultado.xlsx");
    if (xlsxPath == null || !System.IO.File.Exists(xlsxPath))
        throw new FileNotFoundException("No se encontró el XLSX. Usá --xlsx=RUTA para indicar el archivo.");

    xlsxPath = Path.GetFullPath(xlsxPath);
    var sheetName = GetOptionValue(args, "--sheet");
    var filterColumn = GetOptionValue(args, "--filter-column");
    var filterValue = GetOptionValue(args, "--filter-value");
    if ((filterColumn == null) != (filterValue == null))
        throw new ArgumentException("Usá juntos --filter-column y --filter-value.");

    // Por defecto guarda el JSON junto al Excel; si se indicó hoja, incluye su nombre en el archivo.
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

// A partir de aquí comienza el proceso normal: configura el entorno y se autentica contra SharePoint.
var environment = EnvironmentConfig.ResolveFromArgs(args);
var sharePoint = environment.SharePoint;
var authToken = await TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync(
    sharePoint.TenantId,
    sharePoint.ClientId,
    sharePoint.CertificateThumbprint,
    sharePoint.SiteUrl);

using var context = new ClientContext(sharePoint.SiteUrl);
// Agrega el token de acceso a cada solicitud HTTP que haga el cliente de SharePoint.
context.ExecutingWebRequest += (_, request) =>
    request.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authToken;

var dataPath = FindDataFile();
if (dataPath == null)
    throw new FileNotFoundException("No se encontró data.json en el directorio actual ni en sus directorios superiores.");

// El archivo de entrada debe ser un array JSON cuyos elementos incluyan un GUID en la propiedad "id".
using var document = JsonDocument.Parse(await System.IO.File.ReadAllTextAsync(dataPath));
if (document.RootElement.ValueKind != JsonValueKind.Array)
    throw new InvalidDataException("data.json debe contener un array de objetos con la propiedad 'id'.");

// Carga las bibliotecas que se revisarán y ejecuta la consulta pendiente de SharePoint.
var lists = LoadSharePointLibraries(context);
context.ExecuteQuery();
var http = new HttpClient();
var solrManager = new SolrManager(http, environment.Solr, environment.Milvus);
// Esta lista controla los destinos habilitados actualmente; en este caso solo se procesa Solr.
var procesarEnShpSolrMilvus = new[] { "Solr" };
var cambios = new ProcesarCambios(
    http,
    context,
    sharePoint.SiteUrl,
    new LogsManager(http, sharePoint.SiteUrl, "update-errepar-document"),
    new SearchService(context));


foreach (var entry in document.RootElement.EnumerateArray())
{
    // Omite las entradas sin un GUID válido para poder continuar con los demás registros.
    if (!entry.TryGetProperty("id", out var idElement) || !Guid.TryParse(idElement.GetString(), out var guid))
    {
        Console.WriteLine($"ID inválido en data.json: {entry}");
        continue;
    }

    try
    {
        // Busca el documento por su GUID en las bibliotecas configuradas.
        var item = FindItem(lists, guid, context);
        if (item == null)
        {
            Console.WriteLine($"GUID {guid}: no encontrado");
            continue;
        }
        Console.WriteLine($"GUID {guid} | Título: {item["Title"]}");
        // Este bloque opcional agregaría un índice de SharePoint; está comentado y no se ejecuta.
        // var actualizado = cambios.AgregarIndice(item, "4e4b1d54-ac85-473d-a88b-8ccd867af570", "Compraventa");
        // Console.WriteLine(actualizado
        //     ? $"GUID {guid}: índice guardado en SharePoint"
        //     : $"GUID {guid}: sin cambios (índice existente o GUID de término vacío)");

        if (procesarEnShpSolrMilvus.Contains("Solr", StringComparer.OrdinalIgnoreCase))
        {
            // Carga los campos adicionales que necesita el formateador y consulta el elemento en SharePoint.
            context.Load(
                item,
                currentItem => currentItem.ContentType,
                currentItem => currentItem.ParentList,
                currentItem => currentItem.ParentList.Fields);
            context.ExecuteQuery();

            // Convierte el elemento al formato esperado por Solr y lo serializa como un array JSON.
            var itemFormateado = await solrManager.GetItemFormatSync(context, item);
            var payloadSolr = Newtonsoft.Json.JsonConvert.SerializeObject(new[] { itemFormateado });
            var resultadoSolr = await solrManager.EnviarASolr(payloadSolr, item.Id);

            Console.WriteLine(resultadoSolr.ok
                ? $"GUID {guid}: documento actualizado en Solr"
                : $"GUID {guid}: error al actualizar Solr: {resultadoSolr.msg}");

            // El envío a Milvus está preparado como ejemplo, pero permanece desactivado porque está comentado.
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
        // Registra el error de este GUID y continúa con el siguiente elemento del archivo.
        Console.WriteLine($"GUID {guid}: error durante la búsqueda: {exception.Message}");
    }
}


static string? GetOptionValue(string[] arguments, string optionName)
{
    // Reconoce opciones con el formato --opcion=valor y devuelve null si no aparecen.
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
    // Comparte el archivo para lectura aunque otra aplicación lo tenga abierto.
    using var xlsxStream = new FileStream(
        xlsxPath,
        FileMode.Open,
        FileAccess.Read,
        FileShare.ReadWrite | FileShare.Delete);
    using var workbook = new XLWorkbook(xlsxStream);
    // Usa la primera hoja por defecto o busca la hoja solicitada sin distinguir mayúsculas.
    var worksheet = sheetName == null
        ? workbook.Worksheets.FirstOrDefault()
        : workbook.Worksheets.FirstOrDefault(
            sheet => string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase));
    if (worksheet == null)
        throw new InvalidDataException(
            sheetName == null
                ? "El archivo XLSX no contiene hojas."
                : $"No se encontró la hoja '{sheetName}'. Hojas disponibles: {string.Join(", ", workbook.Worksheets.Select(sheet => sheet.Name))}.");

    // Detecta el rango ocupado para no recorrer filas o columnas vacías fuera de la tabla.
    var firstRow = worksheet.FirstRowUsed()?.RowNumber()
        ?? throw new InvalidDataException("La primera hoja del XLSX está vacía.");
    var lastRow = worksheet.LastRowUsed()!.RowNumber();
    var firstColumn = worksheet.FirstColumnUsed()!.ColumnNumber();
    var lastColumn = worksheet.LastColumnUsed()!.ColumnNumber();

    // La primera fila ocupada se interpreta como encabezado de las propiedades JSON.
    var headers = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1)
        .Select(column => worksheet.Cell(firstRow, column).GetString().Trim())
        .ToArray();

    var filterColumnIndex = -1;
    if (filterColumn != null)
    {
        // Resuelve la columna de filtro a partir del encabezado, ignorando diferencias de mayúsculas.
        filterColumnIndex = Array.FindIndex(
            headers,
            header => string.Equals(header, filterColumn, StringComparison.OrdinalIgnoreCase));
        if (filterColumnIndex < 0)
            throw new InvalidDataException(
                $"No se encontró la columna '{filterColumn}' en la hoja '{worksheet.Name}'. Encabezados: {string.Join(" | ", headers)}.");
    }

    // Los encabezados deben existir y ser únicos porque se usan como nombres de propiedades JSON.
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
        // Lee las celdas de la fila; las filas totalmente vacías no generan registros.
        var cells = Enumerable.Range(firstColumn, headers.Length)
            .Select(column => worksheet.Cell(row, column))
            .ToArray();
        if (cells.All(cell => cell.IsEmpty()))
            continue;

        if (filterColumnIndex >= 0
            // El filtro acepta valores que comiencen con el texto indicado, sin distinguir mayúsculas.
            && !cells[filterColumnIndex].GetString().Trim().StartsWith(filterValue!, StringComparison.OrdinalIgnoreCase))
            continue;

        // Asocia cada encabezado con el valor convertido de su celda y agrega el registro al resultado.
        var record = new Dictionary<string, object?>(StringComparer.Ordinal);
        for (var index = 0; index < headers.Length; index++)
            record[headers[index]] = GetJsonCellValue(cells[index]);

        records.Add(record);
    }

    // Genera JSON legible y conserva caracteres como tildes sin escaparlos innecesariamente.
    return JsonSerializer.Serialize(records, new JsonSerializerOptions
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    });
}

static string MakeSafeFileName(string value)
{
    // Reemplaza caracteres no permitidos para que el nombre de hoja pueda usarse como nombre de archivo.
    var invalidCharacters = Path.GetInvalidFileNameChars();
    return string.Join("_", value.Split(invalidCharacters, StringSplitOptions.RemoveEmptyEntries));
}

static object? GetJsonCellValue(IXLCell cell)
{
    // Convierte celdas vacías en null y conserva tipos comunes como número, fecha y booleano.
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

// Busca el archivo de entrada del proceso normal usando la misma estrategia de búsqueda de archivos.
static string? FindDataFile() => FindFile("data.json");

static string? FindFile(string fileName)
{
    // Busca desde el directorio de ejecución y desde el de la aplicación, subiendo por sus directorios padres.
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
    // Nombres de las bibliotecas de SharePoint en las que se buscarán los documentos.
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
        // Prepara cada biblioteca para consultar su título e identificador en la siguiente ejecución de consulta.
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
    // Busca el GUID en cada biblioteca y devuelve el primer elemento coincidente.
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
    // Recorre la biblioteca en páginas de hasta 2000 elementos para evitar cargarla completa de una vez.
    var targetGuid = uniqueId.ToString();
    ListItemCollectionPosition? position = null;

    do
    {
        var query = new CamlQuery
        {
            ListItemCollectionPosition = position,
            // Solicita ID y GUID, ordenados por ID, y habilita la búsqueda recursiva en carpetas.
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
            // SharePoint puede devolver el GUID con llaves; se quitan antes de comparar.
            var itemGuid = candidate["GUID"]?.ToString()?.Trim('{', '}');
            if (!string.Equals(itemGuid, targetGuid, StringComparison.OrdinalIgnoreCase))
                continue;

            // Al encontrarlo, carga todas las propiedades del elemento antes de devolverlo.
            context.Load(candidate);
            context.ExecuteQuery();
            return candidate;
        }

        // Si SharePoint indicó otra página, la consulta siguiente continúa desde esa posición.
        position = items.ListItemCollectionPosition;
    }
    while (position != null);

    return null;
}