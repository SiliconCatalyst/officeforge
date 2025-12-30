# OfficeForge

[![Go Version](https://img.shields.io/github/go-mod/go-version/siliconcatalyst/officeforge)](https://go.dev/)
[![Release](https://img.shields.io/github/v/release/siliconcatalyst/officeforge)](https://github.com/siliconcatalyst/officeforge/releases)
[![Go Report Card](https://goreportcard.com/badge/github.com/siliconcatalyst/officeforge)](https://goreportcard.com/report/github.com/siliconcatalyst/officeforge)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A pure Go library and CLI for generating Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) documents with zero external dependencies.

## Features

- **Zero Dependencies** – Pure Go using only standard library
- **High Performance** – 12,000+ replacements per second
- **Three Formats** – Word, Excel, and PowerPoint support
- **Dual Interface** – Go library or standalone CLI
- **Cross-Platform** – Windows, macOS, Linux
- **Flexible Input** – Keywords work with or without braces

## Performance

**System:** Intel Core i5-8350U @ 1.70GHz, Windows 11, Go 1.23

### Contract Generation Benchmark

```bash
go test -bench=BenchmarkDocxContract -benchmem
```

**Results:**

- **145 documents/second** (9 replacements per document)
- **1,303 replacements/second**
- **6.91ms average per document**
- **282μs per operation** (Go benchmark)
- **9.9KB memory per document**
- **63 allocations per document**

```
BenchmarkDocxContract-8    5427    281999 ns/op    9925 B/op    63 allocs/op
```

### Real-World Test

Processing 1,000 complete contracts (9 fields each) in **under 7 seconds**.

```
╔══════════════════════════════════════════╗
║ Documents processed:     1000            ║
║ Keywords per doc:           9            ║
║ Total replacements:      9000            ║
║ Time taken:              6.91s           ║
║ Throughput:               145 docs/s     ║
╚══════════════════════════════════════════╝
```

### Key Metrics Explained

- **281999 ns/op** (0.28ms) - Time per document operation
- **9925 B/op** (~10KB) - Memory allocated per document
- **63 allocs/op** - Memory allocations (low GC pressure)

### Run Benchmarks Yourself

```bash
cd tests/benchmarks
go test -v -run=TestRealWorldPerformance
go test -bench=. -benchmem
```

## Installation

### CLI Tool

Download from [Releases](https://github.com/siliconcatalyst/officeforge/releases) or install via Go:

```bash
go install github.com/siliconcatalyst/officeforge/cmd/officeforge@latest
```

### Go Library

```bash
go get github.com/siliconcatalyst/officeforge@latest
```

## Quick Start

### CLI Examples

```bash
# Single replacement (Word)
officeforge docx-single -i template.docx -o output.docx -k "CLIENT_NAME" -v "John Doe"

# Multiple replacements (Excel)
officeforge xlsx-multi -i template.xlsx -o output.xlsx -d data.json

# Batch generation (PowerPoint)
officeforge pptx-batch -i template.pptx -o ./output -d records.csv -p "{{NAME}}_slides.pptx"
```

### Library Examples

```go
import (
    "github.com/siliconcatalyst/officeforge/docx"
    "github.com/siliconcatalyst/officeforge/excel"
    "github.com/siliconcatalyst/officeforge/powerpoint"
)

// Word document
docx.ProcessDocxSingle("template.docx", "output.docx", "NAME", "John Doe")

// Excel spreadsheet
excel.ProcessXlsxMulti("template.xlsx", "output.xlsx", map[string]string{
    "TOTAL": "$5,000",
    "DATE": "2024-12-30",
})

// PowerPoint presentation
powerpoint.ProcessPptxBatch("template.pptx", "./output", records, "{{NAME}}_{{INDEX}}.pptx")
```

## CLI Commands

All commands support `-i/--input`, `-o/--output`, `-k/--key`, `-v/--value`, `-d/--data`, `-p/--pattern` flags.

### Word Documents

```bash
docx-single   # Replace one keyword
docx-multi    # Replace multiple keywords from JSON
docx-batch    # Generate multiple documents from CSV/JSON
docx-check    # Verify keywords exist in document
```

### Excel Spreadsheets

```bash
xlsx-single   # Replace one keyword
xlsx-multi    # Replace multiple keywords from JSON
xlsx-batch    # Generate multiple spreadsheets from CSV/JSON
xlsx-check    # Verify keywords exist in spreadsheet
```

### PowerPoint Presentations

```bash
pptx-single   # Replace one keyword
pptx-multi    # Replace multiple keywords from JSON
pptx-batch    # Generate multiple presentations from CSV/JSON
pptx-check    # Verify keywords exist in presentation
```

## Data Formats

### Keywords in Documents

Documents must use double braces:

```
{{CLIENT_NAME}}
{{INVOICE_DATE}}
{{TOTAL_AMOUNT}}
```

### Input Data (Flexible)

Data can be provided **with or without braces**:

**JSON (both formats work):**

```json
{
	"CLIENT_NAME": "John Doe",
	"{{INVOICE_DATE}}": "2024-12-30"
}
```

**CSV (no braces needed):**

```csv
CLIENT_NAME,INVOICE_DATE,AMOUNT
John Doe,2024-12-30,$5000
Jane Smith,2024-12-31,$7500
```

**CLI flags (both work):**

```bash
-k "CLIENT_NAME" -v "John"
-k "{{CLIENT_NAME}}" -v "John"
```

## Batch Processing Patterns

### Sequential Pattern

```bash
-p "document_%d.docx"
# Generates: document_1.docx, document_2.docx, ...
```

### Data-Based Pattern

```bash
-p "{{NAME}}_contract.docx"
# Generates: John_contract.docx, Jane_contract.docx, ...
```

### Combined Pattern

```bash
-p "{{COMPANY}}_{{INDEX}}.docx"
# Generates: Acme_1.docx, Acme_2.docx, ...
```

**Built-in {{INDEX}} placeholder:** Always available in patterns for record numbering.

## Library API

### Word (docx package)

```go
ProcessDocxSingle(inputPath, outputPath, keyword, replacement string) error
ProcessDocxMulti(inputPath, outputPath string, replacements map[string]string) error
ProcessDocxMultipleRecords(inputPath, outputDir string, records []map[string]string, pattern string) error
ProcessDocxMultipleRecordsWithNames(inputPath, outputDir string, records []map[string]string, nameFunc func(map[string]string, int) string) error
```

### Excel (excel package)

```go
ProcessXlsxSingle(inputPath, outputPath, keyword, replacement string) error
ProcessXlsxMulti(inputPath, outputPath string, replacements map[string]string) error
ProcessXlsxMultipleRecords(inputPath, outputDir string, records []map[string]string, pattern string) error
ProcessXlsxMultipleRecordsWithNames(inputPath, outputDir string, records []map[string]string, nameFunc func(map[string]string, int) string) error
```

### PowerPoint (powerpoint package)

```go
ProcessPptxSingle(inputPath, outputPath, keyword, replacement string) error
ProcessPptxMulti(inputPath, outputPath string, replacements map[string]string) error
ProcessPptxMultipleRecords(inputPath, outputDir string, records []map[string]string, pattern string) error
ProcessPptxMultipleRecordsWithNames(inputPath, outputDir string, records []map[string]string, nameFunc func(map[string]string, int) string) error
```

## Integration Examples

### Python

```python
import subprocess, json

data = {"CLIENT_NAME": "John Doe", "AMOUNT": "$5000"}
with open("data.json", "w") as f:
    json.dump(data, f)

subprocess.run(["officeforge", "docx-multi", "-i", "template.docx", "-o", "output.docx", "-d", "data.json"])
```

### Node.js

```javascript
const { execSync } = require("child_process");
execSync(
	'officeforge xlsx-single -i template.xlsx -o output.xlsx -k CLIENT_NAME -v "John Doe"',
);
```

### PHP

```php
shell_exec('officeforge pptx-batch -i template.pptx -o ./output -d records.csv -p "{NAME}_slides.pptx"');
```

## Best Practices

1. **Use descriptive keywords:** `{{CLIENT_FULL_NAME}}` not `{{N}}`
2. **Test templates:** Use `-single` commands to verify keyword placement
3. **Handle errors:** Check exit codes in scripts
4. **Batch processing:** Process large datasets in chunks (100-1000 records)

## Troubleshooting

| Issue                 | Solution                                                |
| --------------------- | ------------------------------------------------------- |
| Keywords not replaced | Ensure exact spelling and `{{BRACES}}` in documents     |
| File not created      | Check output directory exists and has write permissions |
| Memory issues         | Process in smaller batches                              |
| Path errors           | Use absolute paths or check working directory           |

## Development

```bash
git clone https://github.com/siliconcatalyst/officeforge.git
cd officeforge
go mod download
go test ./...
```

## Roadmap

- [x] Word (DOCX) support
- [x] Excel (XLSX) support
- [x] PowerPoint (PPTX) support
- [x] CLI tool with batch processing
- [ ] Image insertion
- [ ] Chart manipulation
- [ ] Package managers (Homebrew, Chocolatey)

## License

MIT License - see LICENSE file

## Support

- [Documentation](https://github.com/siliconcatalyst/officeforge)
- [Issue Tracker](https://github.com/siliconcatalyst/officeforge/issues)

---

Give it a star if you find it useful!
