using System.Text;

internal sealed class ExecutionTextLog : IDisposable
{
    private readonly StreamWriter _writer;
    private readonly object _writeLock = new();

    public string FilePath { get; }

    public ExecutionTextLog(string? itemId = null)
    {
        var logsDirectory = Path.Combine(AppContext.BaseDirectory, "Logs");
        Directory.CreateDirectory(logsDirectory);

        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
        var uniqueId = Guid.NewGuid().ToString("N");
        var itemSuffix = string.IsNullOrWhiteSpace(itemId) ? string.Empty : $"_ID_{itemId}";
        FilePath = Path.Combine(logsDirectory, $"Log_{timestamp}{itemSuffix}_{uniqueId}.txt");
        _writer = new StreamWriter(FilePath, false, new UTF8Encoding(false))
        {
            AutoFlush = true
        };

        Write("INICIO DE EJECUCIÓN");
        if (!string.IsNullOrWhiteSpace(itemId))
            Write($"ID | {itemId}");
    }

    public void Write(string message)
    {
        lock (_writeLock)
            _writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}");
    }

    public void Dispose()
    {
        _writer.Dispose();
    }
}