using System.Diagnostics;
using Nedev.FileConverters.PptxToPpt.Ppt;
using Nedev.FileConverters.PptxToPpt.Pptx;

namespace Nedev.FileConverters.PptxToPpt.Conversion;

public sealed class Converter
{
    public async Task ConvertAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default, bool overwrite = false)
    {
        if (!File.Exists(inputPath))
            throw new FileNotFoundException("Input file not found", inputPath);

        var ext = Path.GetExtension(inputPath).ToLowerInvariant();
        if (ext != ".pptx")
            throw new ArgumentException("Input file must be a .pptx file");

        if (File.Exists(outputPath) && !overwrite)
            throw new IOException($"Output file already exists: {outputPath}. Use overwrite option to replace.");

        var sw = Stopwatch.StartNew();

        var parser = new PptxParser();
        var presentation = await parser.ParseAsync(inputPath);

        var builder = new PptDocumentBuilder();

        foreach (var slide in presentation.Slides)
        {
            builder.AddSlide(slide);
        }

        builder.AddMaster(presentation.MainMaster);
        builder.AddLayouts(presentation.SlideLayouts);

        foreach (var fontName in presentation.Fonts.Keys)
        {
            builder.AddFont(fontName);
        }

        foreach (var media in presentation.MediaFiles)
        {
            var name = media.Key.Substring("ppt/media/".Length);
            builder.AddMedia(name, media.Value);
        }

        await using var outputStream = File.Create(outputPath);
        builder.WriteTo(outputStream);

        sw.Stop();
        Console.WriteLine($"Conversion completed in {sw.ElapsedMilliseconds}ms");
    }

    public void Convert(string inputPath, string outputPath, bool overwrite = false)
    {
        ConvertAsync(inputPath, outputPath, CancellationToken.None, overwrite).GetAwaiter().GetResult();
    }

    public async Task ConvertBatchAsync(IEnumerable<string> inputFiles, string outputDirectory, CancellationToken cancellationToken = default, bool overwrite = false)
    {
        if (!Directory.Exists(outputDirectory))
            Directory.CreateDirectory(outputDirectory);

        var tasks = new List<Task>();
        
        foreach (var inputFile in inputFiles)
        {
            var fileName = Path.GetFileNameWithoutExtension(inputFile);
            var outputPath = Path.Combine(outputDirectory, fileName + ".ppt");
            
            tasks.Add(ConvertAsync(inputFile, outputPath, cancellationToken, overwrite));
        }

        await Task.WhenAll(tasks);
    }

    public void ConvertBatch(IEnumerable<string> inputFiles, string outputDirectory)
    {
        ConvertBatchAsync(inputFiles, outputDirectory).GetAwaiter().GetResult();
    }
}

public sealed class ConverterOptions
{
    public bool Overwrite { get; set; } = false;
    public bool Verbose { get; set; } = false;
    public string? OutputDirectory { get; set; }
}

public sealed class ConverterResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public long ElapsedMilliseconds { get; set; }
    public string? OutputPath { get; set; }
}

public static class ConverterExtensions
{
    public static ConverterResult ConvertWithResult(this Converter converter, string inputPath, string outputPath)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            converter.Convert(inputPath, outputPath);
            sw.Stop();
            return new ConverterResult
            {
                Success = true,
                ElapsedMilliseconds = sw.ElapsedMilliseconds,
                OutputPath = outputPath
            };
        }
        catch (Exception ex)
        {
            sw.Stop();
            return new ConverterResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                ElapsedMilliseconds = sw.ElapsedMilliseconds
            };
        }
    }

    public static async Task<ConverterResult> ConvertWithResultAsync(this Converter converter, string inputPath, string outputPath, CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            await converter.ConvertAsync(inputPath, outputPath, cancellationToken);
            sw.Stop();
            return new ConverterResult
            {
                Success = true,
                ElapsedMilliseconds = sw.ElapsedMilliseconds,
                OutputPath = outputPath
            };
        }
        catch (Exception ex)
        {
            sw.Stop();
            return new ConverterResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                ElapsedMilliseconds = sw.ElapsedMilliseconds
            };
        }
    }
}
