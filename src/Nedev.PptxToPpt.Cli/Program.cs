using Nedev.PptxToPpt.Conversion;

namespace Nedev.PptxToPpt.Cli;

public sealed class Program
{
    public static async Task<int> Main(string[] args)
    {
        if (args.Length == 0)
        {
            PrintUsage();
            return 0;
        }

        var options = new ConverterOptions();
        var inputFiles = new List<string>();

        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg == "-h" || arg == "--help")
            {
                PrintUsage();
                return 0;
            }
            else if (arg == "-o" || arg == "--output")
            {
                if (i + 1 < args.Length)
                {
                    options.OutputDirectory = args[++i];
                }
            }
            else if (arg == "-f" || arg == "--force")
            {
                options.Overwrite = true;
            }
            else if (arg == "-v" || arg == "--verbose")
            {
                options.Verbose = true;
            }
            else if (File.Exists(arg))
            {
                inputFiles.Add(arg);
            }
            else if (Directory.Exists(arg))
            {
                var files = Directory.GetFiles(arg, "*.pptx", SearchOption.AllDirectories);
                inputFiles.AddRange(files);
            }
            else
            {
                Console.WriteLine($"Warning: File or directory not found: {arg}");
            }
        }

        if (inputFiles.Count == 0)
        {
            Console.WriteLine("Error: No input files specified");
            return 1;
        }

        var converter = new Converter();
        var successCount = 0;
        var failCount = 0;

        foreach (var inputFile in inputFiles)
        {
            try
            {
                var fileName = Path.GetFileNameWithoutExtension(inputFile);
                string outputPath;

                if (options.OutputDirectory != null)
                {
                    if (!Directory.Exists(options.OutputDirectory))
                        Directory.CreateDirectory(options.OutputDirectory);
                    outputPath = Path.Combine(options.OutputDirectory, fileName + ".ppt");
                }
                else
                {
                    outputPath = Path.Combine(Path.GetDirectoryName(inputFile) ?? "", fileName + ".ppt");
                }

                if (File.Exists(outputPath) && !options.Overwrite)
                {
                    Console.WriteLine($"Skipping {inputFile}: output file already exists (use -f to overwrite)");
                    failCount++;
                    continue;
                }

                if (options.Verbose)
                    Console.WriteLine($"Converting: {inputFile} -> {outputPath}");

                await converter.ConvertAsync(inputFile, outputPath);
                successCount++;

                if (options.Verbose)
                    Console.WriteLine($"Success: {outputPath}");
            }
            catch (Exception ex)
            {
                failCount++;
                Console.WriteLine($"Error converting {inputFile}: {ex.Message}");
            }
        }

        Console.WriteLine($"\nCompleted: {successCount} succeeded, {failCount} failed");

        return failCount > 0 ? 1 : 0;
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Nedev.PptxToPpt - PPTX to PPT Converter");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  Nedev.PptxToPpt [options] <input files or directories>");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  -o, --output <directory>  Output directory");
        Console.WriteLine("  -f, --force                Overwrite existing files");
        Console.WriteLine("  -v, --verbose              Verbose output");
        Console.WriteLine("  -h, --help                 Show this help");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  Nedev.PptxToPpt file.pptx");
        Console.WriteLine("  Nedev.PptxToPpt -o output file.pptx");
        Console.WriteLine("  Nedev.PptxToPpt -f *.pptx");
        Console.WriteLine("  Nedev.PptxToPpt -o outputdir folder/");
    }
}
