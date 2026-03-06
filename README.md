# Nedev.FileConverters.PptxToPpt

PPTX to PPT converter - convert PowerPoint Open XML (.pptx) to legacy binary format (.ppt) with high performance. This library also integrates with the `Nedev.FileConverters.Core` NuGet package: a file‑converter adapter is registered automatically so that the core's static converter/registry can be used by other projects.

This project has no third‑party dependencies other than the Core package.

## Project Status: Functional Prototype (Current)

This project is currently in early development. It provides a functional foundation for converting PPTX files to the legacy PPT format. While core features like text and basic shapes are supported, advanced formatting and complex elements are still in development.

## Features

- **Pure .NET 10 Implementation**: No third-party library dependencies (no Interop, no OpenXML SDK).
- **Custom CFBF Writer**: Built-in implementation of the Compound File Binary Format (MS-CFB).
- **Core Conversion**: Basic support for slides, text, and common shapes.
- **Media Support**: Preliminary support for embedded images.
- **CLI Tool**: Batch conversion support via command line.
- **Performance Focused**: Minimal memory footprint and high-speed processing.

## Supported Elements

- [x] Slide parsing and generation
- [x] Text boxes and basic text formatting
- [x] Simple geometric shapes (Rectangles, Ellipses, etc.)
- [x] Embedded Pictures (JPEG, PNG)
- [x] Slide Notes
- [x] Basic Master Slide support
- [x] Theme/Font mapping

## Roadmap

- [ ] **Advanced Formatting**: Complex text effects, paragraph styles, and bullet points.
- [ ] **Complex Shapes**: SmartArt, group shapes (nested), and advanced Bezier curves.
- [ ] **Tables & Charts**: Full support for PowerPoint tables and OLE-embedded charts.
- [ ] **Transitions & Animations**: Mapping PPTX entrance/exit effects to PPT equivalents.
- [ ] **Audio/Video**: Support for embedded multimedia streams.
- [ ] **Encryption**: Support for password-protected documents.

## Project Structure

```
src/
├── Nedev.FileConverters.PptxToPpt/           # Core library (depends on Nedev.FileConverters.Core NuGet package)
│   ├── Cff/                    # Compound File Format writer (MS-CFB)
│   ├── Ppt/                    # PPT binary format generator (MS-PPT)
│   ├── Pptx/                   # PPTX parser (OpenXML)
│   └── Conversion/             # High-level conversion orchestrator
└── Nedev.FileConverters.PptxToPpt.Cli/        # Command-line interface application
```

## CLI Usage

```bash
# Convert single file
dotnet run --project src/Nedev.FileConverters.PptxToPpt.Cli input.pptx  # uses Nedev.FileConverters.Core NuGet package for logging/exception

# Specify output directory
dotnet run --project src/Nedev.FileConverters.PptxToPpt.Cli -o output input.pptx

# Batch convert
dotnet run --project src/Nedev.FileConverters.PptxToPpt.Cli -f *.pptx
```

## Library Usage

This project produces a nuget package (`Nedev.FileConverters.PptxToPpt`) containing
both `net8.0` and `netstandard2.1` builds. To consume in another project:

1. Add a package reference:
    ```xml
    <PackageReference Include="Nedev.FileConverters.PptxToPpt" Version="0.1.0" />
    ```
   or run `dotnet add package Nedev.FileConverters.PptxToPpt --version 0.1.0`.

2. Call the static converter exposed by the package:
    ```csharp
    using (var input = File.OpenRead("presentation.pptx"))
    using (var output = Nedev.FileConverters.Converter.Convert(input, "pptx", "ppt"))
    using (var fs = File.Create("presentation.ppt"))
    {
        output.CopyTo(fs);
    }
    ```

3. **Implementing `IFileConverter`**

   The core package defines an interface:
   ```csharp
   namespace Nedev.FileConverters.Core {
       public interface IFileConverter {
           Stream Convert(Stream input);
       }
   }
   ```
   You can implement this in your own library to register new conversions. A sample
   adapter already exists in this repo (`Conversion/PptxToPptFileConverter.cs`):
   it writes the incoming stream to temporary files and delegates to the internal
   converter logic.

   To register a converter, use the DI extension method provided by the core
   package:
   ```csharp
   var services = new ServiceCollection();
   services.AddFileConverter("pptx","ppt", new PptxToPptFileConverter());
   var provider = services.BuildServiceProvider();
   ```
   After registration the converter becomes available through
   `Nedev.FileConverters.Converter.Convert(...)` or the `ConverterRegistry`.

   You may decorate your implementation with the `[FileConverter]` attribute
   to allow automatic discovery when the package scans assemblies.

---

The CLI application demonstrates both the basic usage and how the adapter is
registered on startup; refer to its `Program.cs` for a working example.

## Build

Built with the new .NET 10 features.

```bash
dotnet build src/Nedev.FileConverters.PptxToPpt.slnx
```

## Requirements

- .NET 10.0 SDK
