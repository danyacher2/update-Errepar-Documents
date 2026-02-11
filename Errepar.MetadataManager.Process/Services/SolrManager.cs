using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Errepar.MetadataManager.Process.Services
{
    // Esqueleto de cliente para Solr. Ajusta URL y esquema de campos según tu core/index.
    public class SolrManager
    {
        private readonly HttpClient _http;
        private readonly string _solrBaseUrl; // p. ej. http://localhost:8983/solr/yourcore

        public SolrManager(HttpClient httpClient, string solrBaseUrl)
        {
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _solrBaseUrl = solrBaseUrl?.TrimEnd('/') ?? throw new ArgumentNullException(nameof(solrBaseUrl));
        }

        public async Task IndexAsync(IEnumerable<object> documents, CancellationToken cancellationToken = default)
        {
            // POST a /update?commit=true con JSON de documentos
            var url = $"{_solrBaseUrl}/update?commit=true";
            await _http.PostAsJsonAsync(url, documents, cancellationToken).ConfigureAwait(false);
        }

        public async Task DeleteByIdAsync(IEnumerable<string> ids, CancellationToken cancellationToken = default)
        {
            var url = $"{_solrBaseUrl}/update?commit=true";
            var deletePayload = new { delete = ids };
            await _http.PostAsJsonAsync(url, deletePayload, cancellationToken).ConfigureAwait(false);
        }

        public async Task<string> SearchRawAsync(string q, int rows = 10, CancellationToken cancellationToken = default)
        {
            var url = $"{_solrBaseUrl}/select?q={Uri.EscapeDataString(q)}&rows={rows}&wt=json";
            var resp = await _http.GetStringAsync(url, cancellationToken).ConfigureAwait(false);
            return resp;
        }
    }
}