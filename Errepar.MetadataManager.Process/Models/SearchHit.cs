using System;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using System;
using System.Linq;
using System.Threading;

namespace Errepar.MetadataManager.Process.Models
{
    public class SearchHit
    {
        public string Path { get; set; }
    }
    public class SearchService
    {
        private readonly ClientContext _context;

        public SearchService(ClientContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public SearchHit ObtenerElementoPorGuid(string shpItemGUID)
        {
            int batchSize = 500;
            int startRow = 0;

            while (true)
            {
                KeywordQuery keywordQuery = new KeywordQuery(_context)
                {
                    QueryText = shpItemGUID,
                    RowLimit = batchSize,
                    StartRow = startRow
                };

                SearchExecutor searchExecutor = new SearchExecutor(_context);
                ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);

                ExecuteQueryRetry(_context);

                var table = results.Value.FirstOrDefault(t => t.TableType == "RelevantResults");
                if (table == null || table.ResultRows == null || !table.ResultRows.Any())
                    return null;

                foreach (var row in table.ResultRows)
                {
                    if (row.ContainsKey("OriginalPath"))
                    {
                        return new SearchHit
                        {
                            Path = row["OriginalPath"]?.ToString() ?? ""
                        };
                    }
                }

                startRow += batchSize;
            }
        }

        private static void ExecuteQueryRetry(ClientContext context, int retryCount = 3, int delayMs = 1000)
        {
            int attempts = 0;

            while (true)
            {
                try
                {
                    context.ExecuteQuery();
                    return;
                }
                catch (ClientRequestException)
                {
                    attempts++;
                    if (attempts > retryCount) throw;
                    Thread.Sleep(delayMs);
                }
                catch (ServerException)
                {
                    attempts++;
                    if (attempts > retryCount) throw;
                    Thread.Sleep(delayMs);
                }
            }
        }
    }
}
