# Nedev.FileConverters.PptxToPpt

[![Build](https://img.shields.io/badge/build-passing-green)](#)
[![License](https://img.shields.io/badge/license-MIT-blue)](#)

A high‑performance, **pure .NET** converter that transforms PowerPoint Open XML (`.pptx`) documents into the legacy binary `.ppt` format. The library is designed for integration with the [`Nedev.FileConverters.Core`](https://www.nuget.org/packages/Nedev.FileConverters.Core) package; an adapter is automatically registered so that the core's global converter registry can be leveraged by host applications.

- ✅ No third‑party dependencies (no Office Interop, no OpenXML SDK)
- ✅ Cross‑targeted to `net8.0` and `netstandard2.1`

Licensed under the [MIT License](LICENSE).

---

## Project Status

**Functional prototype / early development.**

The converter already handles basic slides, text, images, and shapes, and it now includes run‑level rich‑text formatting (fonts, size, bold, italic). Many advanced PPTX features remain on the roadmap, but the core conversion pipeline is stable and can be used in production scenarios that do not rely on exotic effects.

## Key Features

- 🔧 **Self‑contained .NET** – written entirely in C#; only dependency is `Nedev.FileConverters.Core`.
- 📄 **CFBF Writer** – a homegrown implementation of Microsoft’s Compound File Binary Format (MS‑CFB) enables creation of `.ppt` containers.
- 🔄 **Conversion Engine** – parses `.pptx` files (Open XML) and generates binary PPT slides, text boxes, and simple shapes.
- 🎨 **Rich Text Formatting** – run‑level support for font, size, bold, italic, underline and color; paragraph styles include basic alignment/level; bullet detection (character or auto‑number) prepends text.
- 🖼️ **Image Embedding** – handles JPEG/PNG slides pictures.
- 🖥️ **CLI Application** – command‑line wrapper for batch conversions.
- ⚡ **Performance‑oriented** – small memory footprint and fast processing suitable for server environments.

## Supported Elements

The converter currently understands and reproduces the following PPTX constructs:

- ✅ Slide markup and layout
- ✅ Text boxes with run‑level font/size/bold/italic/underline/color and simple bullets (character or auto-number)
- ✅ Primitive shapes (rectangle, ellipse, line, etc.)
- ✅ Grouped shapes with basic translation (now with scale and rotation encoded)
- ✅ Embedded images (JPEG/PNG)
- ✅ Slide notes
- ✅ Simple master slide information
- ✅ Theme and font mapping

> **Note:** Other elements (charts, tables, SmartArt, animations, etc.) are not yet supported.

## Roadmap

The following enhancements are planned for future releases:

1. **Advanced text formatting** – paragraph styles and numbering (basic bullets and colors/underlines now supported).
2. **Complex shapes & SmartArt** – groups (including better scaling/rotation semantics), Bézier paths, and custom geometry.
3. **Tables & Charts** – full fidelity conversion of PPTX tables and embedded Office charts.
4. **Animations & Transitions** – approximate effects in binary PPT.
5. **Audio/Video** – embed multimedia streams and control metadata.
6. **Document protection** – decryption/encryption of password‑protected files.

Contributions and feature requests are welcome via GitHub issues.

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

## Command‑Line Interface (CLI)

The companion CLI project provides a simple batch converter. Example commands:

```bash
# convert a single file
dotnet run --project src/Nedev.FileConverters.PptxToPpt.Cli input.pptx

# write output to specific folder
dotnet run --project src/Nedev.FileConverters.PptxToPpt.Cli -o output input.pptx

# wildcards for multiple files
dotnet run --project src/Nedev.FileConverters.PptxToPpt.Cli -f "*.pptx"
```

The CLI registers the adapter with `Nedev.FileConverters.Core` so the same API can be used
programmatically; inspect `Program.cs` for a minimal example.


## Library Usage

The core project compiles to a NuGet package (`Nedev.FileConverters.PptxToPpt`) with targets
for `net8.0` and `netstandard2.1`.

### Installing

```bash
dotnet add package Nedev.FileConverters.PptxToPpt --version 0.1.0
```

### Basic API

```csharp
using var input = File.OpenRead("presentation.pptx");
using var output = Nedev.FileConverters.Converter.Convert(input, "pptx", "ppt");
using var fs     = File.Create("presentation.ppt");
output.CopyTo(fs);
```

The static `Converter` class comes from `Nedev.FileConverters.Core` and automatically
invokes any registered adapters (including the one provided in this repo).


---

The CLI application demonstrates both the basic usage and how the adapter is
registered on startup; refer to its `Program.cs` for a working example.

## Building the Solution

Requirements: **.NET 10.0 SDK**

From the repository root run:

```bash
dotnet build src/Nedev.FileConverters.PptxToPpt.slnx
```

The solution includes the core library, CLI, and test projects.

## Requirements

- .NET 10.0 SDK (see global.json for exact version)

---

For detailed design notes and contribution guidelines, see the `docs/` folder
(coming soon) or open an issue on GitHub.
