using Errepar.MetadataManager.Process.Models;
using Microsoft.SharePoint.Client;
using System;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Errepar.MetadataManager.Process.Services
{
    /// <summary>
    /// Procesa los cambios de metadata para los items de SharePoint.
    /// </summary>
   
public class ProcesarCambios
    {
        private readonly HttpClient _http;
        private readonly ClientContext _context;
        private readonly string _siteUrl;
        private readonly LogsManager _logsManager;
        private readonly SearchService _searchService;

        public ProcesarCambios(
            HttpClient httpClient, 
            ClientContext context, 
            string siteUrl, 
            LogsManager logsManager,
            SearchService searchService)
        {
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _siteUrl = siteUrl?.TrimEnd('/') ?? throw new ArgumentNullException(nameof(siteUrl));
            _logsManager = logsManager ?? throw new ArgumentNullException(nameof(logsManager));
            _searchService = searchService ?? throw new ArgumentNullException(nameof(searchService));
        }

        public void Procesar(ListItem sharepointItem, JsonElement cambio)
        {
            var action = GetString(cambio, "action");
            var campo = GetString(cambio, "campo");
            var tipo = GetString(cambio, "tipo");
            var valor = GetString(cambio, "valor");

            if (string.IsNullOrWhiteSpace(action) || string.IsNullOrWhiteSpace(campo))
            {
                Console.WriteLine("⚠ Cambio inválido (faltan campos obligatorios)");
                return;
            }

            switch (action.ToLower())
            {
                case "agregar":
                    Agregar(sharepointItem, campo, tipo, valor, cambio);
                    break;

                case "quitar":
                    Quitar(sharepointItem, campo, tipo, valor, cambio);
                    break;

                case "reemplazar":
                    Reemplazar(sharepointItem, campo, tipo, valor, cambio);
                    break;

                default:
                    Console.WriteLine($"⚠ Acción desconocida: {action}");
                    break;
            }
        }
        private string GetString(JsonElement el, string prop)
        {
            if (el.TryGetProperty(prop, out var p))
                return p.ToString();

            return null;
        }
        public void Agregar(ListItem item, string campo, string tipo, string valor, JsonElement cambio)
        {
            tipo = tipo?.ToLower();

            switch (tipo)
            {
                // 🔹 Campos simples
                case "text":
                case "multitext":
                case "numbers":
                case "date":                    
                    var internalName = campo;

                    var convertedValue = ConvertirValor(tipo, valor);

                    if (convertedValue == null)
                    {
                        return;
                    }


                    item[internalName] = convertedValue;
                    break;

                case "boolean":
                    if (bool.TryParse(valor, out var b))
                        item[campo] = b;
                    break;

                // 🔹 Choice
                case "choice":
                    item[campo] = valor;
                    break;

                // 🔹 Lookup
                case "lookup":
                    {
                        var lookupId = int.Parse(valor); // si viene id
                        item[campo] = new FieldLookupValue { LookupId = lookupId };
                    }
                    break;

                // 🔹 Managed Metadata
                case "metadata":
                    {
                        // Esperado: "Label|TermGuid"
                        var parts = valor.Split('|');
                        if (parts.Length == 2)
                        {
                            item[campo] = parts[0] + "|" + parts[1];
                        }
                    }
                    break;

                default:
                    Console.WriteLine($"⚠ Tipo no soportado en AGREGAR: {tipo}");
                    break;
            }
        }

        public void Quitar(ListItem item, string campo, string tipo, string valor, JsonElement cambio)
        {
            tipo = tipo?.ToLower();

            if (!item.FieldValues.ContainsKey(campo))
                return;

            switch (tipo)
            {
                case "text":
                case "multitext":
                    {
                        var actual = item[campo]?.ToString() ?? "";
                        if (actual.Contains(valor))
                        {
                            actual = actual.Replace(valor, "").Trim();
                            item[campo] = actual;
                        }
                    }
                    break;

                case "numbers":
                case "date":
                case "boolean":
                    item[campo] = null;
                    break;

                case "choice":
                    item[campo] = null;
                    break;

                case "lookup":
                    item[campo] = null;
                    break;

                case "metadata":
                    item[campo] = null;
                    break;

                default:
                    Console.WriteLine($"⚠ Tipo no soportado en QUITAR: {tipo}");
                    break;
            }
        }
        public void Reemplazar(ListItem item, string campo, string tipo, string valor, JsonElement cambio)
        {
            Quitar(item, campo, tipo, valor, cambio);
            Agregar(item, campo, tipo, valor, cambio);
        }

        object ConvertirValor(string tipo, string valor)
        {
            if (valor == null) return null;

            tipo = tipo?.ToLower();

            switch (tipo)
            {
                case "text":
                case "multitext":
                    return valor;

                case "numbers":
                    if (double.TryParse(valor, out var num))
                        return num;
                    return null;

                case "date":
                    if (DateTime.TryParse(valor, out var dt))
                        return dt;
                    return null;

                case "boolean":
                    if (bool.TryParse(valor, out var b))
                        return b;
                    return null;

                case "choice":
                    return valor; // string exacto del choice

                case "lookup":
                    // valor = ID del item lookup
                    if (int.TryParse(valor, out var id))
                        return new FieldLookupValue { LookupId = id };
                    return null;

                case "metadata":
                    // formato: Label|TermGuid
                    var parts = valor.Split('|');
                    if (parts.Length == 2)
                        return parts[0] + "|" + parts[1];
                    return null;

                default:
                    return null;
            }
        }
    }
}
