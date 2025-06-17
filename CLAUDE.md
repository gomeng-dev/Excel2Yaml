# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Memories

- VSTO 프로젝트는 내가 직접 빌드하고 테스트 해야하니까 테스트는 네가 직접 하려고 하지 말아줘 , 대신 나에게 테스트 제안을 해줘
- 항상 한국어로 답변하고 한국어로 생각해줘
- 하드코딩을 하지 말고 구조적으로 코딩할 수 있도록 해줘
- 구조적 설계를 1순위로 생각해 하드코딩을 배제하고 당장 비슷한 결과가 나오는게 아니라 범용적인 사용이 가능하도록 로직을 설계해
- 빌드는 항상 내가 직접 할게
- 새로운 파일을 추가하면 항상 ExcelToYamlAddin.csproj에 반영해줘
- 리팩토링 작업을 할 때는 Docs/Excel2Yaml_리팩토링_계획서.md 를 참고하고 업데이트 하면서 작업해줘

## Build Commands

### Build the Project
```bash
# Build in Release mode (recommended)
msbuild ExcelToYamlAddin.sln /p:Configuration=Release

# Build in Debug mode
msbuild ExcelToYamlAddin.sln /p:Configuration=Debug

# Clean and rebuild
msbuild ExcelToYamlAddin.sln /t:Clean,Build /p:Configuration=Release
```

### Create Deployment Package
```powershell
# Build and create deployment package
.\Setup\build-and-deploy.ps1

# Skip build if already built
.\Setup\build-and-deploy.ps1 -SkipBuild

# Create ZIP package
.\Setup\build-and-deploy.ps1 -CreateZip
```

### Install the Add-in
```powershell
# Install the add-in (run as Administrator)
.\Setup\install.ps1

# Uninstall the add-in
.\Setup\install.ps1 -Uninstall
```

## Code Architecture

### Core Components Overview

The Excel2YAML add-in follows a VSTO (Visual Studio Tools for Office) architecture with a clear separation of concerns:

1. **Entry Points**
   - `ThisAddIn.cs`: VSTO add-in initialization and lifecycle management
   - `Ribbon.cs`: UI controls and user interaction handling
   - `RibbonUI.xml`: Ribbon UI definition

2. **Core Conversion Pipeline**
   - `ExcelReader.cs`: Main conversion orchestrator that handles file I/O
   - `Core/SchemeParser.cs`: Parses Excel schema structure (rows with $ markers)
   - `Core/SchemeNode.cs`: Represents schema structure nodes (MAP, ARRAY, PROPERTY, etc.)
   - `Core/YamlGenerator.cs`: Converts parsed schema to YAML format
   - `Core/Generator.cs`: Converts parsed schema to JSON format

3. **Post-Processing**
   - `Core/YamlPostProcessors/YamlMergeKeyPathsProcessor.cs`: Merges YAML entries by ID
   - `Core/YamlPostProcessors/YamlFlowStyleProcessor.cs`: Applies flow style formatting
   - `Core/YamlPostProcessors/YamlToJsonProcessor.cs`: YAML to JSON conversion
   - `Core/YamlPostProcessors/YamlToXmlConverter.cs`: YAML to XML conversion

4. **Configuration Management**
   - `Config/SheetPathManager.cs`: Manages per-sheet output paths and settings
   - `Config/ExcelToYamlConfig.cs`: Global conversion configuration
   - `Core/ExcelConfigManager.cs`: Excel-based configuration storage

5. **Utilities**
   - `Core/OrderedYamlFactory.cs`: Maintains property order in YAML output
   - `Core/OrderedJsonFactory.cs`: Maintains property order in JSON output
   - `Core/ExcelCellValueResolver.cs`: Handles Excel cell value type detection

### Key Data Flow

1. User clicks conversion button in Ribbon
2. `Ribbon.cs` validates sheets (marked with `!` prefix) and prepares conversion
3. `ExcelReader.cs` reads the Excel file using ClosedXML
4. `SchemeParser.cs` analyzes schema structure (looking for `$scheme_end` marker)
5. `SchemeNode.cs` builds hierarchical representation of the schema
6. `YamlGenerator.cs` or `Generator.cs` converts to YAML/JSON
7. Post-processors apply additional transformations
8. Files are saved to configured paths

### Schema Markers

The Excel schema uses special markers to define structure:
- `$[]`: Array container
- `${}`: Object/Map container  
- `$key`: Dynamic key from cell value
- `$value`: Dynamic value from cell value
- `^`: Ignore this cell
- `$scheme_end`: Marks end of schema definition (must be in merged cell spanning all columns)

### Important Design Decisions

1. **Sheet Naming**: Only sheets with `!` prefix are converted
2. **Property Order**: Both YAML and JSON maintain Excel column order
3. **Configuration Storage**: Settings stored in hidden Excel sheet `_ExcelToYamlConfig`
4. **Empty Fields**: Controlled by both global and per-sheet settings
5. **Post-Processing**: Applied after initial conversion for merge/flow operations

## Testing Guidelines

The project doesn't have automated tests. Manual testing should cover:
1. Basic schema types (arrays, objects, nested structures)
2. Post-processing features (merge, flow style)
3. Different output formats (YAML, JSON, XML)
4. Sheet-specific configurations
5. Error cases (missing schema end marker, invalid structure)

## Common Issues and Solutions

1. **Schema Not Recognized**: Ensure `$scheme_end` is in the last row, merged across all columns
2. **Sheet Not Converting**: Check sheet name starts with `!`
3. **Empty Output**: Verify sheet is enabled in settings and has valid output path
4. **Post-Processing Not Applied**: Check configuration in hidden `_ExcelToYamlConfig` sheet