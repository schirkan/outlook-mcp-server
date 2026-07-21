using Microsoft.Extensions.Logging;

namespace OutlookMcpServer;

/// <summary>
/// Minimaler dateibasierter Logger-Provider. Schreibt Logeintraege zeilenweise
/// in eine konfigurierte Datei (Append-Modus). Erweiterbar fuer Rotation, falls
/// noetig.
/// <para>
/// Aktivierung: <c>appsettings.json</c> &rarr; <c>Logging:FilePath</c> oder
/// ENV <c>OUTLOOKMCPSERVER_Logging__FilePath</c>.
/// </para>
/// </summary>
public sealed class FileLoggerProvider : ILoggerProvider
{
    private readonly string _filePath;
    private readonly object _lock = new();

    public FileLoggerProvider(string filePath)
    {
        _filePath = filePath;
    }

    public ILogger CreateLogger(string categoryName) => new FileLogger(this, categoryName);

    public void Dispose()
    {
        // Kein unmanaged Resource zu schliessen — StreamWriter wird pro Eintrag geoeffnet.
    }

    private void WriteLine(string line)
    {
        lock (_lock)
        {
            try
            {
                File.AppendAllText(_filePath, line + Environment.NewLine);
            }
            catch
            {
                // File-Logger darf niemals den eigentlichen Tool-Aufruf abbrechen.
                // Im Worst-Case bleibt die Datei gesperrt oder die Platte voll —
                // wir verwerfen stillschweigend (stderr-Console-Logger laeuft parallel).
            }
        }
    }

    private sealed class FileLogger : ILogger
    {
        private readonly FileLoggerProvider _provider;
        private readonly string _category;

        public FileLogger(FileLoggerProvider provider, string category)
        {
            _provider = provider;
            _category = category;
        }

        public IDisposable BeginScope<TState>(TState state) where TState : notnull => NullScope.Instance;
        public bool IsEnabled(LogLevel logLevel) => logLevel >= LogLevel.Information;

        public void Log<TState>(
            LogLevel logLevel,
            EventId eventId,
            TState state,
            Exception? exception,
            Func<TState, Exception?, string> formatter)
        {
            if (!IsEnabled(logLevel)) return;
            var timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            var message = formatter(state, exception);
            var line = $"{timestamp} [{logLevel}] {_category}: {message}";
            if (exception is not null)
            {
                line += Environment.NewLine + exception;
            }
            _provider.WriteLine(line);
        }

        private sealed class NullScope : IDisposable
        {
            public static readonly NullScope Instance = new();
            public void Dispose() { }
        }
    }
}