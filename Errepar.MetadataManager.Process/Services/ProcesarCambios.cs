using Errepar.MetadataManager.Process.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using static System.Net.WebRequestMethods;

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

        public void Procesar(ListItem sharepointItem, JsonElement cambio, int itemId)
            {
            var action = GetString(cambio, "action");
            var campo = GetString(cambio, "campo");
            var tipo = GetString(cambio, "tipo");
            var valor = GetString(cambio, "valor");
            var vuevoValor = GetString(cambio, "nuevoValor");

            string taxonomyGuid = null;
            string taxonomyName = null;
            string taxonomyNewName = null;
            string taxonomyNewGuid = null;

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
            if (cambio.TryGetProperty("taxonomyReplaceTerms", out taxonomyTermsElement)
                && taxonomyTermsElement.ValueKind == JsonValueKind.Array)
                {
                foreach (var term in taxonomyTermsElement.EnumerateArray())
                    {
                    taxonomyNewGuid = term.GetProperty("id").GetString();
                    taxonomyNewName = term.GetProperty("labels")[0].GetProperty("name").GetString();
                    break; // single taxonomy
                    }
                }

            switch (action.ToLower())
                {
                case "agregar":
                    Agregar(sharepointItem, campo, tipo, valor, taxonomyName, taxonomyGuid, itemId);
                    break;

                case "quitar":
                    Quitar(sharepointItem, campo, tipo, valor, taxonomyName, taxonomyGuid, itemId);
                    break;

                case "reemplazar":
                    Reemplazar(sharepointItem, campo, tipo, valor, vuevoValor, taxonomyName, taxonomyGuid, taxonomyNewName, taxonomyNewGuid, itemId);
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
        public async void Agregar(ListItem item, string campo, string tipo, string valor, string taxName = "", string taxGuid = "", int itemId = 0)
            {
            tipo = tipo?.ToLower();

            if (!item.FieldValues.ContainsKey(campo))
                {
                Console.WriteLine($"⚠ El campo '{campo}' no existe en la lista");
                await _logsManager.SaveLogErrorCambios(itemId, new ItemLogError
                    {
                    ItemId = item.Id.ToString(),
                    Fecha = DateTime.Now,
                    Mensaje = $"El Campo {campo} no exite en este item",
                    Estado = "Error en ProcesarCambios"
                    });

                return;
                }
            try
                {
                switch (tipo)
                    {
                    // 🔹 Campos simples
                    case "numbers":
                    case "date":
                    case "text":
                    case "multitext":
                            {
                            if (item[campo] != valor)
                                {
                                item[campo] = valor;
                                }
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
                    case "metadata":
                            {
                            if (string.IsNullOrWhiteSpace(taxGuid))
                                return;

                            _context.Load(item);
                            _context.ExecuteQuery();

                            var field = item.ParentList.Fields.GetByInternalNameOrTitle(campo);
                            _context.Load(field);
                            _context.ExecuteQuery();

                            var taxField = _context.CastTo<TaxonomyField>(field);

                            if (taxField.AllowMultipleValues)
                                {
                                var termStrings = new List<string>();

                                if (item[campo] != null)
                                    {
                                    var currentCollection = item[campo] as TaxonomyFieldValueCollection;

                                    if (currentCollection != null)
                                        {
                                        foreach (var t in currentCollection)
                                            {
                                            if (t.TermGuid.Equals(taxGuid, StringComparison.OrdinalIgnoreCase))
                                                return;

                                            termStrings.Add($"-1;#{t.Label}|{t.TermGuid}");
                                            }
                                        }
                                    }

                                var nombreFinal = taxName.Contains(":")
                                    ? taxName.Split(':').Last()
                                    : taxName;

                                termStrings.Add($"-1;#{nombreFinal}|{taxGuid}");

                                string nuevoValorInterno = string.Join(";#", termStrings);

                                var nuevaCollection = new TaxonomyFieldValueCollection(
                                    _context,
                                    nuevoValorInterno,
                                    taxField);

                                taxField.SetFieldValueByValueCollection(item, nuevaCollection);
                                }
                            else
                                {
                                var taxValue = new TaxonomyFieldValue
                                    {
                                    Label = taxName,
                                    TermGuid = taxGuid,
                                    WssId = -1
                                    };

                                taxField.SetFieldValueByValue(item, taxValue);
                                }

                            item.Update();
                            _context.ExecuteQuery();
                            }
                        break;

                    default:
                        Console.WriteLine($"⚠ Tipo no soportado en AGREGAR: {tipo}");
                        break;
                    }
                }
            catch (Exception ex)
                {
                await _logsManager.SaveLogErrorCambios(item.Id, new ItemLogError
                    {
                    ItemId = item.Id.ToString(),
                    Fecha = DateTime.Now,
                    Mensaje = ex.ToString(),
                    Estado = "Error en ProcesarCambios"
                    });

                Console.Error.WriteLine("Excepción no controlada: " + ex);
                }
            }

        public async void Quitar(ListItem item, string campo, string tipo, string valor, string taxName, string taxGuid, int itemId = 0)
            {
            tipo = tipo?.ToLower();

            if (!item.FieldValues.ContainsKey(campo))
                {
                Console.WriteLine($"⚠ El campo '{campo}' no existe en la lista {item.ParentList.Title}");
                await _logsManager.SaveLogErrorCambios(itemId, new ItemLogError
                    {
                    ItemId = item.Id.ToString(),
                    Fecha = DateTime.Now,
                    Mensaje = $"El Campo {campo} no exite en este item",
                    Estado = "Error en ProcesarCambios"
                    });
                return;
                }

            if (!item.FieldValues.ContainsKey(campo))
                return;
            try
                {

                switch (tipo)
                    {
                    case "text":
                    case "date":
                    case "multitext":
                    case "numbers":
                            {
                            if (item[campo] == null)
                                return;
                            if (item[campo] == valor)
                                {
                                item[campo] = null;
                                }

                            }
                        break;

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
                            if (string.IsNullOrWhiteSpace(taxGuid))
                                return;

                            var field = item.ParentList.Fields.GetByInternalNameOrTitle(campo);
                            _context.Load(field);
                            _context.ExecuteQuery();

                            var taxField = _context.CastTo<TaxonomyField>(field);

                            _context.Load(item);
                            _context.ExecuteQuery();

                            if (taxField.AllowMultipleValues)
                                {
                                var currentValues = item[campo] as TaxonomyFieldValueCollection;

                                if (currentValues == null || currentValues.Count == 0)
                                    return;

                                // Filtrar términos
                                var remainingTerms = currentValues
                                    .Cast<TaxonomyFieldValue>()
                                    .Where(t => !t.TermGuid.Equals(taxGuid, StringComparison.OrdinalIgnoreCase))
                                    .ToList();

                                // Construir string en formato interno
                                string newValue = string.Join(";#", remainingTerms.Select(t =>
                                    $"-1;#{t.Label}|{t.TermGuid}"));

                                // Crear nueva colección
                                var newCollection = new TaxonomyFieldValueCollection(_context, newValue, taxField);

                                taxField.SetFieldValueByValueCollection(item, newCollection);
                                item.Update();
                                _context.ExecuteQuery();
                                }
                            else
                                {
                                // SINGLE VALUE
                                var currentResult = taxField.GetFieldValueAsTaxonomyFieldValue(campo);
                                _context.ExecuteQuery();

                                var currentValue = currentResult.Value;

                                if (currentValue != null &&
                                    currentValue.TermGuid.Equals(taxGuid, StringComparison.OrdinalIgnoreCase))
                                    {
                                    item[campo] = null;
                                    }
                                }

                            item.Update();
                            _context.ExecuteQuery();
                            }
                        break;

                    default:
                        Console.WriteLine($"⚠ Tipo no soportado en QUITAR: {tipo}");
                        break;
                    }
                }
            catch (Exception ex)
                {
                await _logsManager.SaveLogErrorCambios(item.Id, new ItemLogError
                    {
                    ItemId = item.Id.ToString(),
                    Fecha = DateTime.Now,
                    Mensaje = ex.ToString(),
                    Estado = "Error en ProcesarCambios"
                    });

                Console.Error.WriteLine("Excepción no controlada: " + ex);
                }
            }

        private bool EsCampoRequerido(Field field)
            {
            _context.Load(field, f => f.Required);
            _context.ExecuteQuery();
            return field.Required;
            }
        public async void Reemplazar(ListItem item, string campo, string tipo, string valor, string vuevoValor, string taxName = "", string taxGuid = "", string taxNewName = "", string taxNewGuid = "", int itemId = 0)
            {
            tipo = tipo?.ToLower();

            var fieldValidation = item.ParentList.Fields.GetByInternalNameOrTitle(campo);
            if (!item.FieldValues.ContainsKey(campo))
                {
                Console.WriteLine($"⚠ El campo '{campo}' no existe en la lista {item}");
                await _logsManager.SaveLogErrorCambios(itemId, new ItemLogError
                    {
                    ItemId = item.Id.ToString(),
                    Fecha = DateTime.Now,
                    Mensaje = $"El Campo {campo} no exite en este item",
                    Estado = "Error en ProcesarCambios"
                    });
                return;
                }

            try
                {

                switch (tipo)
                    {
                    case "text":
                    case "multitext":
                    case "numbers":
                            {
                            //Console.WriteLine(item[campo].ToString() +"=="+ new Date(valor));
                            if (item[campo]?.ToString() == valor)
                                item[campo] = vuevoValor;
                            }
                        break;
                    case "date":
                            {
                            var fechaItem = (DateTime)item[campo];
                            var fechaValor = DateTime.Parse(valor, null, DateTimeStyles.RoundtripKind);

                            if (fechaItem.Date == fechaValor.Date)
                                {
                                item[campo] = vuevoValor;

                                }
                            }
                        break;

                    case "boolean":
                            {
                            var actual = item[campo];
                            if (actual == null)
                                return;

                            if (actual.ToString().Equals(valor, StringComparison.OrdinalIgnoreCase))
                                {
                                item[campo] = vuevoValor;
                                }
                            }
                        break;

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

                                if (actualValue is string[] arr)
                                    valores.AddRange(arr);
                                else
                                    valores.AddRange(
                                        actualValue.ToString()
                                                   .Split(new[] { ";#" }, StringSplitOptions.RemoveEmptyEntries)
                                    );

                                bool cambiado = false;

                                for (int i = 0; i < valores.Count; i++)
                                    {
                                    if (valores[i].Equals(valor, StringComparison.OrdinalIgnoreCase))
                                        {
                                        valores[i] = vuevoValor;
                                        cambiado = true;
                                        }
                                    }

                                if (cambiado)
                                    item[campo] = valores.ToArray();
                                }
                            else
                                {
                                // 🔹 SINGLE CHOICE
                                if (actualValue.ToString().Equals(vuevoValor, StringComparison.OrdinalIgnoreCase))
                                    {
                                    item[campo] = vuevoValor;
                                    }
                                }
                            }
                        break;

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

                    case "metadata":
                            {
                            if (string.IsNullOrWhiteSpace(taxGuid))
                                return;

                            var field = item.ParentList.Fields.GetByInternalNameOrTitle(campo);
                            _context.Load(field);
                            _context.ExecuteQuery();

                            var taxField = _context.CastTo<TaxonomyField>(field);

                            _context.Load(item);
                            _context.ExecuteQuery();

                            if (taxField.AllowMultipleValues)
                                {
                                var currentValues = item[campo] as TaxonomyFieldValueCollection;

                                if (currentValues == null || currentValues.Count == 0)
                                    return;

                                var remainingTerms = currentValues
                                    .Cast<TaxonomyFieldValue>()
                                    .Where(t => !t.TermGuid.Equals(taxGuid, StringComparison.OrdinalIgnoreCase))
                                    .ToList();

                                var nuevoTerm = new TaxonomyFieldValue
                                    {
                                    Label = taxNewName,
                                    TermGuid = taxNewGuid,
                                    WssId = -1
                                    };

                                remainingTerms.Add(nuevoTerm);

                                string newValue = string.Join(";#", remainingTerms.Select(t =>
                                    $"-1;#{t.Label}|{t.TermGuid}"));

                                var newCollection = new TaxonomyFieldValueCollection(_context, newValue, taxField);

                                taxField.SetFieldValueByValueCollection(item, newCollection);
                                item.Update();
                                _context.ExecuteQuery();
                                }
                            else
                                {
                                var currentResult = taxField.GetFieldValueAsTaxonomyFieldValue(item[campo].ToString());
                                _context.ExecuteQuery();

                                var currentValue = currentResult.Value;

                                if (currentValue != null &&
                                    currentValue.TermGuid.Equals(taxGuid, StringComparison.OrdinalIgnoreCase))
                                    {
                                    var nuevoTerm = new TaxonomyFieldValue
                                        {
                                        Label = taxNewName,
                                        TermGuid = taxNewGuid,
                                        WssId = -1
                                        };

                                    taxField.SetFieldValueByValue(item, nuevoTerm);
                                    }
                                }

                            item.Update();
                            _context.ExecuteQuery();

                            }
                        break;

                    default:
                        Console.WriteLine($"⚠ Tipo no soportado en QUITAR: {tipo}");
                        break;
                    }
                }
            catch (Exception ex)
                {
                await _logsManager.SaveLogErrorCambios(item.Id, new ItemLogError
                    {
                    ItemId = item.Id.ToString(),
                    Fecha = DateTime.Now,
                    Mensaje = ex.ToString(),
                    Estado = "Error en ProcesarCambios"
                    });

                Console.Error.WriteLine("Excepción no controlada: " + ex);
                }
            }


        }
    }
