using Errepar.MetadataManager.Process.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
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

            string taxonomyGuid = null;
            string taxonomyName = null;

            JsonElement taxonomyTermsElement;
            if (cambio.TryGetProperty("taxonomyTerms", out taxonomyTermsElement)
                && taxonomyTermsElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var term in taxonomyTermsElement.EnumerateArray())
                {
                    taxonomyGuid = term.GetProperty("id").GetString();
                    taxonomyName = term.GetProperty("labels")[0].GetProperty("name").GetString();
                    break; // single taxonomy
                }
            }

            switch (action.ToLower())
            {
                case "agregar":
                    Agregar(sharepointItem, campo, tipo, valor, taxonomyName, taxonomyGuid);
                    break;

                case "quitar":
                    Quitar(sharepointItem, campo, tipo, valor, taxonomyName, taxonomyGuid);
                    break;

                case "reemplazar":
                    // Reemplazar(sharepointItem, campo, tipo, valor, taxonomyName, taxonomyGuid);
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
        public void Agregar(ListItem item, string campo, string tipo, string valor, string taxName = "", string taxGuid = "")
        {
            tipo = tipo?.ToLower();

            switch (tipo)
            {
                // 🔹 Campos simples
                case "numbers":
                    {
                        var actual = item[campo];

                        double numeroActual = 0;

                        if (actual != null && double.TryParse(actual.ToString(), out var parsed))
                            numeroActual = parsed;

                        var nuevoValor = numeroActual + valor;

                        item[campo] = nuevoValor;
                    }
                    break;
                case "date": {
                        if (item[campo] == null)
                        {
                            item[campo] = valor;
                        }
                    }
                    break;
                case "text":
                case "multitext":                
                
                    {
                        var internalName = campo;
                        var convertedValue = ConvertirValor(tipo, valor);

                        if (convertedValue == null)
                            return;

                        var actual = item[internalName];

                        // Si no existe valor previo → set directo
                        if (actual == null || string.IsNullOrWhiteSpace(actual.ToString()))
                        {
                            item[internalName] = convertedValue.ToString();
                            return;
                        }

                        var textoActual = actual.ToString();
                        var textoNuevo = convertedValue.ToString();

                        // 🧠 Evitar duplicados exactos
                        if (textoActual.Contains(textoNuevo))
                            return;

                        // 👉 Concatenar con separador
                        item[internalName] = $"{textoActual} {textoNuevo}";
                    }
                    break;

                case "boolean":
                    {
                        if (!bool.TryParse(valor, out var b))
                            return;

                        var actual = item[campo];

                        if (actual != null && actual is bool currentBool && currentBool == b)
                            return;

                        item[campo] = b;
                    }
                    break;

                // 🔹 Choice
                case "choice":
                    var spField = item.ParentList.Fields.GetByInternalNameOrTitle(campo);
                    _context.Load(spField);
                    _context.ExecuteQuery();

                    var actualValue = item[campo];

                    if (spField.TypeAsString == "MultiChoice")
                    {
                        var valores = new List<string>();

                        if (actualValue != null)
                        {
                            if (actualValue is string[])
                            {
                                valores.AddRange((string[])actualValue);
                            }
                            else
                            {
                                valores.AddRange(
                                    actualValue.ToString()
                                               .Split(new[] { ";#" }, StringSplitOptions.RemoveEmptyEntries)
                                );
                            }
                        }

                        if (!valores.Contains(valor))
                            valores.Add(valor);

                        // ✅ ASIGNACIÓN CORRECTA
                        item[campo] = valores.ToArray();   // string[]
                    }
                    else
                    {
                        item[campo] = valor;
                    }
                    break;

                // 🔹 Lookup
                case "lookup":
                    {
                        var lookupId = int.Parse(valor);
                        var actual = item[campo] as FieldLookupValue;

                        if (actual != null && actual.LookupId == lookupId)
                            return;

                        item[campo] = new FieldLookupValue { LookupId = lookupId };
                    }
                    break;

                // 🔹 Managed Metadata (single value)
                case "metadata":
                    {
                        if (string.IsNullOrWhiteSpace(taxGuid))
                            return;

                        var field = item.ParentList.Fields.GetByInternalNameOrTitle(campo);
                        _context.Load(field);
                        _context.ExecuteQuery();

                        var taxField = _context.CastTo<TaxonomyField>(field);

                        // 🔍 Leer valor actual
                        var currentResult = taxField.GetFieldValueAsTaxonomyFieldValue(item[campo]?.ToString());
                        _context.ExecuteQuery();

                        var currentValue = currentResult.Value; // 👈 ahora sí

                        // 🧠 Si ya existe y es el mismo término → no tocar
                        if (currentValue != null &&
                            currentValue.TermGuid != null &&
                            currentValue.TermGuid.Equals(taxGuid, StringComparison.OrdinalIgnoreCase))
                        {
                            return;
                        }

                        var taxValue = new TaxonomyFieldValue
                        {
                            Label = taxName,
                            TermGuid = taxGuid,
                            WssId = -1
                        };

                        taxField.SetFieldValueByValue(item, taxValue);
                    }
                    break;

                default:
                    Console.WriteLine($"⚠ Tipo no soportado en AGREGAR: {tipo}");
                    break;
            }
        }
        public void Quitar(ListItem item, string campo, string tipo, string valor, string taxName, string taxGuid)
        {
            tipo = tipo?.ToLower();

            if (!item.FieldValues.ContainsKey(campo))
                return;

            switch (tipo)
            {
                // 🔹 TEXT / MULTITEXT
                case "text":
                case "multitext":
                    {
                        var actual = item[campo]?.ToString() ?? "";

                        if (string.IsNullOrWhiteSpace(actual))
                            return;

                        if (!actual.Contains(valor))
                            return; 

                        var nuevo = actual.Replace(valor, "").Trim();

                        
                        nuevo = nuevo.Replace(";;", ";").Trim(';').Trim();

                        item[campo] = string.IsNullOrWhiteSpace(nuevo) ? null : nuevo;
                    }
                    break;

                // 🔹 NUMBERS / DATE / BOOLEAN
                case "numbers":
                    {
                        if (item[campo] == null)
                            return;

                        var actualStr = item[campo].ToString();

                        if (!actualStr.Contains(valor))
                            return;

                        var nuevoStr = actualStr.Replace(valor, "").Trim();

                        // Si queda vacío → null
                        if (string.IsNullOrWhiteSpace(nuevoStr))
                        {
                            item[campo] = null;
                            return;
                        }

                        // Validar que siga siendo número
                        if (double.TryParse(nuevoStr, out var nuevoNumero))
                            item[campo] = nuevoNumero;
                        else
                            item[campo] = null; // o lanzar error según lógica
                    }
                    break;
                case "date":
                case "boolean":
                    {
                        var actual = item[campo];
                        if (actual == null)
                            return;

                        // si coincide → quitar
                        if (actual.ToString() == valor)
                            item[campo] = null;
                    }
                    break;

                // 🔹 CHOICE
                case "choice":
                    {
                        var spField = item.ParentList.Fields.GetByInternalNameOrTitle(campo);
                        _context.Load(spField);
                        _context.ExecuteQuery();

                        var actualValue = item[campo];

                        if (actualValue == null)
                            return;

                        // 🔹 MULTI CHOICE
                        if (spField.TypeAsString == "MultiChoice")
                        {
                            var valores = new List<string>();

                            if (actualValue is string[])
                            {
                                valores.AddRange((string[])actualValue);
                            }
                            else
                            {
                                valores.AddRange(
                                    actualValue.ToString()
                                               .Split(new[] { ";#" }, StringSplitOptions.RemoveEmptyEntries)
                                );
                            }

                            // Quitar valor
                            valores.RemoveAll(v => v == valor);

                            // Guardar
                            item[campo] = valores.Count == 0 ? null : valores.ToArray();
                        }
                        else
                        {
                            // 🔹 SINGLE CHOICE
                            if (actualValue.ToString() == valor)
                                item[campo] = null;
                        }
                    }
                    break;

                // 🔹 LOOKUP (single o multi)
                case "lookup":
                    {
                        if (item[campo] == null)
                            return;

                        // multi lookup
                        if (item[campo] is FieldLookupValue[] multi)
                        {
                            var nuevaLista = multi
                                .Where(l => l.LookupId.ToString() != valor)
                                .ToArray();

                            if (nuevaLista.Length != multi.Length)
                                item[campo] = nuevaLista.Length == 0 ? null : nuevaLista;
                        }
                        // single lookup
                        else if (item[campo] is FieldLookupValue single)
                        {
                            if (single.LookupId.ToString() == valor)
                                item[campo] = null;
                        }
                    }
                    break;

                // 🔹 METADATA (single o multi)
                case "metadata":
                    {
                        var field = item.ParentList.Fields.GetByInternalNameOrTitle(campo);
                        _context.Load(field);
                        _context.ExecuteQuery();

                        var taxField = _context.CastTo<TaxonomyField>(field);

                        // obtener valores actuales
                        var currentValue = item[campo]?.ToString();

                        if (string.IsNullOrWhiteSpace(currentValue))
                            return;

                        // multi metadata
                        if (currentValue.Contains(";"))
                        {
                            var valores = currentValue.Split(';').ToList();

                            // formato: Label|Guid
                            var nuevo = valores
                                .Where(v => !v.EndsWith("|" + taxGuid, StringComparison.OrdinalIgnoreCase))
                                .ToList();

                            if (nuevo.Count == valores.Count)
                                return; // ❌ no existía → no hace nada

                            if (nuevo.Count == 0)
                            {
                                // ⚠ verificar required
                                if (!EsCampoRequerido(field))
                                    taxField.SetFieldValueByValue(item, null);
                            }
                            else
                            {
                                var final = string.Join(";", nuevo);
                                var col = taxField.GetFieldValueAsTaxonomyFieldValueCollection(final);
                                taxField.SetFieldValueByValueCollection(item, col);
                            }
                        }
                        else
                        {
                            // single metadata
                            if (currentValue.EndsWith("|" + taxGuid, StringComparison.OrdinalIgnoreCase))
                            {
                                if (!EsCampoRequerido(field))
                                    taxField.SetFieldValueByValue(item, null);
                            }
                        }
                    }
                    break;

                default:
                    Console.WriteLine($"⚠ Tipo no soportado en QUITAR: {tipo}");
                    break;
            }
        }

        private bool EsCampoRequerido(Field field)
        {
            _context.Load(field, f => f.Required);
            _context.ExecuteQuery();
            return field.Required;
        }
        //public void Reemplazar(ListItem item, string campo, string tipo, string valor, JsonElement cambio)
        //{
        //    Quitar(item, campo, tipo, valor, cambio);
        //    Agregar(item, campo, tipo, valor, cambio);
        //}

        object ConvertirValor(string tipo, string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
                return null;

            tipo = tipo?.ToLower();

            switch (tipo)
            {
                case "text":
                case "multitext":
                case "choice":
                    return valor;

                case "numbers":
                    if (int.TryParse(valor, out var i))
                        return i;
                    if (double.TryParse(valor, out var d))
                        return d;
                    return null;

                case "date":
                    if (DateTime.TryParse(valor, out var dt))
                        return dt;
                    return null;

                case "boolean":
                    if (bool.TryParse(valor, out var b))
                        return b;
                    return null;

                case "lookup":
                    if (int.TryParse(valor, out var id))
                        return new FieldLookupValue { LookupId = id };
                    return null;

                case "metadata":
                    // ⚠️ Este método NO debe usarse para metadata
                    // metadata se maneja por TaxonomyField.SetFieldValueByValue()
                    return null;

                default:
                    return null;
            }
        }
    }
}
