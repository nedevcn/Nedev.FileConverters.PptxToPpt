# Nedev.PptxToPpt

PPTX to PPT converter - convert PowerPoint Open XML (.pptx) to legacy binary format (.ppt) with high performance and zero third-party dependencies.

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
├── Nedev.PptxToPpt/           # Core library
│   ├── Cff/                    # Compound File Format writer (MS-CFB)
│   ├── Ppt/                    # PPT binary format generator (MS-PPT)
│   ├── Pptx/                   # PPTX parser (OpenXML)
│   └── Conversion/             # High-level conversion orchestrator
└── Nedev.PptxToPpt.Cli/        # Command-line interface application
```

## Usage

```bash
# Convert single file
dotnet run --project src/Nedev.PptxToPpt.Cli input.pptx

# Specify output directory
dotnet run --project src/Nedev.PptxToPpt.Cli -o output input.pptx

# Batch convert
dotnet run --project src/Nedev.PptxToPpt.Cli -f *.pptx
```

## Build

Built with the new .NET 10 features.

```bash
dotnet build Nedev.PptxToPpt.slnx
```

## Requirements

- .NET 10.0 SDK
