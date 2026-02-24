using Errepar.MetadataManager.Process.Auth;
using Errepar.MetadataManager.Process.Models;
using Errepar.MetadataManager.Process.Services;
using Errepar.MetadataManager.Process.Services;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.News.DataModel;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
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
    string clientId = "8688eed4-7464-4288-9820-34849fd19296";

    // Autenticación con certificado usando thumbprint
    // Obtén el thumbprint desde el certificado instalado en Windows (ver pasos más abajo)
    string certificateThumbprint = Environment.GetEnvironmentVariable("CERT_THUMBPRINT") 
        ?? "452079A2697BC9646023FAE02876488654BBDB2C"; // Reemplaza con tu thumbprint real

    // Generar token con certificado (thumbprint)
    var authToken = await TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync( tenantId, clientId, certificateThumbprint, siteUrl);
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

    Console.WriteLine($"Items obtenidos: {items.Count}");
    foreach (var it in items)
    {
        Console.WriteLine($"Id={it.Id} | Título='{it.Titulo}' | Estado='{it.EstadoProceso}' | EjecutadoPor='{it.EjecutadoPor}' | Link='{it.Link}'");
        foreach (var activo in it.Activos)
        { 
            var hit = searchService.ObtenerElementoPorGuid(activo);

            //if (hit != null)
            //{
                //Console.WriteLine("Path encontrado: " + hit.Path);

                var siteUri = new Uri(siteUrl);
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
                    var file = context.Web.GetFileByServerRelativeUrl("/sites/ErreparDesarrollo/Documento/00-sin-obra/2024/20240126103422795/20240126103422795.html");

                    context.Load(file, f => f.ListItemAllFields, f => f.Exists);
                    context.ExecuteQuery();

                    if (!file.Exists)
                    {
                        Console.WriteLine("❌ El path no corresponde a un archivo");
                        return;
                    }

                    var item = file.ListItemAllFields;

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

                    var folder = Path.Combine(AppContext.BaseDirectory, "searchLogs");
                    Directory.CreateDirectory(folder);

                    var filePath = Path.Combine(folder, $"ListItems/Search_{activo}.json");
                    System.IO.File.WriteAllText(filePath, json, Encoding.UTF8);
                    Console.WriteLine($"✅ JSON completo guardado en: {filePath}");

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
                ItemListId = (int)item.FieldValues["Id"],
                LibraryName = item.FieldValues["Title"].ToString(),
                TimeStamp = (int)item.FieldValues["TimeStamp"],
                UrlItem = item.FieldValues["FileRef"].ToString(),
                Activo = activo,
                Procesado = true,
                Fecha = DateTime.Now
            };

            await logsManager.SaveLogActivosProcesados(it.Id, new[] { logActivo });
        }


        await logsManager.SyncLogs(it.Id,  cts.Token);

    }
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