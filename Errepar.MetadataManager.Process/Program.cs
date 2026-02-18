using Errepar.MetadataManager.Process.Auth;
using Errepar.MetadataManager.Process.Services;
using Microsoft.SharePoint.Client;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

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
    var authToken = await TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync(
        tenantId, clientId, certificateThumbprint, siteUrl);

    using var http = new HttpClient();
    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authToken);

    var sp = new SharePointManager(http, siteUrl, "Metadata Manager", authToken);
    using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(2));

    var items = await sp.GetPendingOrPausedAsync(cts.Token).ConfigureAwait(false);

    Console.WriteLine($"Items obtenidos: {items.Count}");
    foreach (var it in items)
    {
        Console.WriteLine($"Id={it.Id} | Título='{it.Titulo}' | Estado='{it.EstadoProceso}' | EjecutadoPor='{it.EjecutadoPor}' | Link='{it.Link}'");
    }

  
    /* ========== EJEMPLO: Usar funciones estáticas para actualizar archivos (borrar después) ==========*/

    // Ejemplo 1: Actualizar un archivo de texto agregando contenido
    if (items.Count > 0)
    {
        var itemId = items[0].Id;
        var fileName = "log.txt";

        // Primera vez: crear el archivo
        //var contenido1 = System.Text.Encoding.UTF8.GetBytes("Línea 1: Inicio del proceso\n");
        //var url1 = await SharePointAttachmentHelper.UpdateOrCreateAttachmentAsync(
        //    http, siteUrl, "Metadata Manager", itemId, fileName, contenido1, appendMode: true, cts.Token);
        //Console.WriteLine($"Archivo creado: {url1}");

        // Segunda vez: agregar más contenido al mismo archivo
        var contenido2 = System.Text.Encoding.UTF8.GetBytes("Línea 2: Proceso completado\n");
        var url2 = await SharePointAttachmentHelper.UpdateOrCreateAttachmentAsync(
            http, siteUrl, "Metadata Manager", itemId, fileName, contenido2, appendMode: true, cts.Token);
        Console.WriteLine($"Contenido agregado: {url2}");

        // El archivo ahora contiene ambas líneas
    }

    // Ejemplo 2: Verificar si existe un archivo y leer su contenido
    if (items.Count > 0)
    {
        var itemId = items[0].Id;
        var fileName = "log.txt";

        bool existe = await SharePointAttachmentHelper.AttachmentExistsAsync(
            http, siteUrl, "Metadata Manager", itemId, fileName, cts.Token);
        Console.WriteLine($"¿Archivo '{fileName}' existe? {existe}");

        if (existe)
        {
            var contenido = await SharePointAttachmentHelper.GetAttachmentContentAsync(
                http, siteUrl, "Metadata Manager", itemId, fileName, cts.Token);

            if (contenido != null)
            {
                var texto = System.Text.Encoding.UTF8.GetString(contenido);
                Console.WriteLine($"Contenido del archivo:\n{texto}");
            }
        }
    }

    // Ejemplo 3: Reemplazar completamente un archivo (appendMode: false)
    if (items.Count > 0)
    {
        var itemId = items[0].Id;
        var fileName = "config.json";

        var nuevoConfig = System.Text.Encoding.UTF8.GetBytes("{\"version\": \"2.0\", \"enabled\": true}");
        var url = await SharePointAttachmentHelper.UpdateOrCreateAttachmentAsync(
            http, siteUrl, "Metadata Manager", itemId, fileName, nuevoConfig, appendMode: false, cts.Token);
        Console.WriteLine($"Archivo reemplazado: {url}");
    }

    // Ejemplo 4: Agregar datos binarios (ej: imagen)
    if (items.Count > 0 && System.IO.File.Exists("C:\\temp\\imagen.jpg"))
    {
        var itemId = items[0].Id;
        var fileName = "foto.jpg";

        var imagenBytes = await System.IO.File.ReadAllBytesAsync("C:\\temp\\imagen.jpg", cts.Token);
        var url = await SharePointAttachmentHelper.UpdateOrCreateAttachmentAsync(
            http, siteUrl, "Metadata Manager", itemId, fileName, imagenBytes, appendMode: false, cts.Token);
        Console.WriteLine($"Imagen subida: {url}");
    }

   /*========== FIN: Ejemplo funciones estáticas ==========*/
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