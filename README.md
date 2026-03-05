# Nedev.PptxToPpt

PPTX to PPT converter - convert PowerPoint Open XML (.pptx) to legacy binary format (.ppt)

## Features

- Pure .NET implementation - no third-party dependencies
- Convert PPTX to PPT format
- Supports batch conversion
- Command-line interface

## Project Structure

```
src/
├── Nedev.PptxToPpt/           # Core library
│   ├── Cff/                    # Compound File Format writer
│   ├── Ppt/                    # PPT binary format generator
│   ├── Pptx/                   # PPTX parser
│   └── Conversion/             # Conversion logic
└── Nedev.PptxToPpt.Cli/        # CLI application
```

## Usage

```bash
# Convert single file
dotnet run --project src/Nedev.PptxToPpt.Cli input.pptx

# Specify output directory
dotnet run --project src/Nedev.PptxToPpt.Cli -o output input.pptx

# Batch convert
dotnet run --project src/Nedev.PptxToPpt.Cli -f *.pptx

# Help
dotnet run --project src/Nedev.PptxToPpt.Cli -- --help
```

## Options

- `-o, --output <directory>` - Output directory
- `-f, --force` - Overwrite existing files
- `-v, --verbose` - Verbose output
- `-h, --help` - Show help

## Build

```bash
dotnet build Nedev.PptxToPpt.slnx
```

## Requirements

- .NET 10.0 SDK
