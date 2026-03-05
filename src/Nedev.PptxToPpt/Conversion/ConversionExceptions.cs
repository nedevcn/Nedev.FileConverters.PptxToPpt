namespace Nedev.PptxToPpt.Conversion;

public sealed class ConversionException : Exception
{
    public string FilePath { get; }

    public ConversionException(string message) : base(message) { FilePath = ""; }

    public ConversionException(string message, Exception innerException) : base(message, innerException) { FilePath = ""; }

    public ConversionException(string message, string filePath) : base(message)
    {
        FilePath = filePath;
    }

    public ConversionException(string message, string filePath, Exception innerException) : base(message, innerException)
    {
        FilePath = filePath;
    }
}

public interface ILogger
{
    void Info(string message);
    void Warning(string message);
    void Error(string message, Exception? ex = null);
    void Debug(string message);
}

public sealed class ConsoleLogger : ILogger
{
    private readonly bool _verbose;

    public ConsoleLogger(bool verbose = false)
    {
        _verbose = verbose;
    }

    public void Info(string message)
    {
        Console.WriteLine($"[INFO] {message}");
    }

    public void Warning(string message)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"[WARN] {message}");
        Console.ResetColor();
    }

    public void Error(string message, Exception? ex = null)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"[ERROR] {message}");
        if (ex != null)
        {
            Console.WriteLine($"  Exception: {ex.Message}");
        }
        Console.ResetColor();
    }

    public void Debug(string message)
    {
        if (_verbose)
        {
            Console.WriteLine($"[DEBUG] {message}");
        }
    }
}

public sealed class NullLogger : ILogger
{
    public static readonly NullLogger Instance = new();

    private NullLogger() { }

    public void Info(string message) { }
    public void Warning(string message) { }
    public void Error(string message, Exception? ex = null) { }
    public void Debug(string message) { }
}
