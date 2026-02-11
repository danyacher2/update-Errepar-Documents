using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Errepar.MetadataManager.Process.Services
{
    // Cliente mínimo para Milvus REST/SDK. Normalmente Milvus se consume con su SDK (gRPC/protobuf).
    // Aquí se deja un esqueleto con llamadas HTTP para ilustrar integración vía REST proxy si existiera.
    public class MilvusManager
    {
        private readonly HttpClient _http;
        private readonly string _milvusBaseUrl;

        public MilvusManager(HttpClient httpClient, string milvusBaseUrl)
        {
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _milvusBaseUrl = milvusBaseUrl?.TrimEnd('/') ?? throw new ArgumentNullException(nameof(milvusBaseUrl));
        }

        public async Task CreateCollectionAsync(object collectionSpec, CancellationToken cancellationToken = default)
        {
            var url = $"{_milvusBaseUrl}/collections";
            await _http.PostAsJsonAsync(url, collectionSpec, cancellationToken).ConfigureAwait(false);
        }

        public async Task InsertVectorsAsync(object insertPayload, CancellationToken cancellationToken = default)
        {
            var url = $"{_milvusBaseUrl}/vectors/insert";
            await _http.PostAsJsonAsync(url, insertPayload, cancellationToken).ConfigureAwait(false);
        }

        public async Task<string> SearchAsync(object searchPayload, CancellationToken cancellationToken = default)
        {
            var url = $"{_milvusBaseUrl}/vectors/search";
            var resp = await _http.PostAsJsonAsync(url, searchPayload, cancellationToken).ConfigureAwait(false);
            return await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        }

        public async Task DeleteAsync(object deletePayload, CancellationToken cancellationToken = default)
        {
            var url = $"{_milvusBaseUrl}/vectors/delete";
            await _http.PostAsJsonAsync(url, deletePayload, cancellationToken).ConfigureAwait(false);
        }
    }
}