using Microsoft.Office.SharePoint.Tools;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.News.DataModel;
using System;
using System;
using System.Linq;
using System.Text;
using System.Text.Json;
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
            int totalProcesados = 0;

            var resultadosJson = new List<Dictionary<string, object>>();

            while (true)
            {
                KeywordQuery keywordQuery = new KeywordQuery(_context)
                {
                    QueryText = shpItemGUID,
                    //QueryText = $"Site:\"https://erreparsa.sharepoint.com/sites/ErreparDesarrollo\"",
                    RowLimit = batchSize,
                    StartRow = startRow
                };

                SearchExecutor searchExecutor = new SearchExecutor(_context);
                ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);

                ExecuteQueryRetry(_context);

                var table = results.Value.FirstOrDefault(t => t.TableType == "RelevantResults");

                if (table == null || table.ResultRows == null || !table.ResultRows.Any())
                {
                    //Console.WriteLine($"✅ Fin del procesamiento. Total procesados: {totalProcesados}");
                    break; // ✅ sale del while, no return
                }

                int batchCount = table.ResultRows.Count();
                totalProcesados += batchCount;

                //Console.WriteLine($"📦 Procesando batch desde {startRow} - batch: {batchCount} - total: {totalProcesados}");

                foreach (var row in table.ResultRows)
                {
                    var rowData = new Dictionary<string, object>();

                    foreach (var key in row.Keys)
                    {
                        rowData[key] = row[key];
                    }

                    resultadosJson.Add(rowData);
                }

                startRow += batchSize;

                var json = JsonSerializer.Serialize(resultadosJson, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                var folder = Path.Combine(AppContext.BaseDirectory, "searchLogs");
                Directory.CreateDirectory(folder);

                var filePath = Path.Combine(folder, $"Search_{shpItemGUID}.json");
                System.IO.File.WriteAllText(filePath, json, Encoding.UTF8);

            }

            // 🔥 Guardar JSON UNA SOLA VEZ (al final)


            Console.WriteLine($"📊 Total registros: {resultadosJson.Count}");

            // 🔁 Retorno funcional
            if (resultadosJson.Count > 0 && resultadosJson[0].ContainsKey("OriginalPath"))
            {
                return new SearchHit
                {
                    Path = resultadosJson[0]["OriginalPath"]?.ToString() ?? ""
                };
            }

            return null;
        }






        //public SearchHit ObtenerElementoPorGuid(string shpItemGUID)
        //{
        //    int batchSize = 500;
        //    int startRow = 0;

        //    while (true)
        //    {
        //        KeywordQuery keywordQuery = new KeywordQuery(_context)
        //        {
        //            QueryText = "ley",
        //            RowLimit = batchSize,
        //            StartRow = startRow
        //        };

        //        SearchExecutor searchExecutor = new SearchExecutor(_context);
        //        ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);

        //        ExecuteQueryRetry(_context);

        //        var table = results.Value.FirstOrDefault(t => t.TableType == "RelevantResults");
        //        if (table == null || table.ResultRows == null || !table.ResultRows.Any())
        //            return null;

        //        foreach (var row in table.ResultRows)
        //        {
        //            if (row.ContainsKey("OriginalPath"))
        //            {
        //                return new SearchHit
        //                {
        //                    Path = row["OriginalPath"]?.ToString() ?? ""
        //                };
        //            }
        //        }

        //        startRow += batchSize;
        //    }
        //}

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
