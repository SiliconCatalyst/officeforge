# OfficeForge

**OfficeForge** – A pure Go library for generating Word, Excel, and PowerPoint documents with zero external dependencies. Built on the standard library for maximum portability, security, and control.

# Features

-   Create and manipulate Word, Excel, and PowerPoint files
-   Zero external libraries — just `zip`, `xml`, and `bytes`
-   Fast, portable, and secure
-   Ideal for server-side automation or static document generation
-   Boosts adminstrative tasks by automating the creation of documents that rely on external data (Client information, Statistics, Data, Graphs)

# Installation

```bash
go get github.com/siliconcatalyst/officeforge
```

# Usage

This library provides two main functions for processing DOCX files with keyword replacements:

## Import

```go
import (
    "github.com/siliconcatalyst/officeforge/docx"
)
```

## Function 1: ProcessDocxSingle

Creates a single output DOCX file with one keyword-replacement pair applied.

```go
err := docx.ProcessDocxSingle("contract.docx", "contract_john.docx", "CLIENT_NAME", "John Smith")
```

### Parameters:

-   `inputPath` (string): Path to the source DOCX file that contains the keywords to be replaced
-   `outputPath` (string): Path where the new DOCX file with replacements will be saved
-   `keyword` (string): The exact text string to search for in the document
-   `replacement` (string): The text that will replace all instances of the keyword

### Use Cases:

-   Creating personalized documents from templates
-   Generating single contracts or letters with specific details
-   Quick one-off document customization
-   Testing individual replacements before batch processing

### Example Scenarios:

```go
// Generate a personalized contract
docx.ProcessDocxSingle("template_contract.docx", "smith_contract.docx", "{{CLIENT}}", "John Smith")

// Create a customized letter
docx.ProcessDocxSingle("letter_template.docx", "welcome_letter.docx", "COMPANY_NAME", "Tech Solutions Inc.")

// Replace dates in documents
docx.ProcessDocxSingle("report.docx", "monthly_report.docx", "REPORT_DATE", "July 2024")
```

## Function 2: ProcessDocxMulti

Creates a single output DOCX file with multiple keyword-replacement pairs applied simultaneously.

```go
batchReplacements := map[string]string{
    "CLIENT_NAME":     "John Smith",
    "COMPANY_NAME":    "Smith Industries",
    "CONTRACT_DATE":   "2024-07-16",
    "CONTRACT_AMOUNT": "$5,000",
    "PROJECT_NAME":    "Website Development",
    "DEADLINE":        "2024-08-30",
}
err := docx.ProcessDocxMulti("contract_template.docx", "completed_contract.docx", batchReplacements)
```

### Parameters:

-   `inputPath` (string): Path to the source DOCX template file
-   `outputPath` (string): Path where the final DOCX file with all replacements will be saved
-   `replacements` (map[string]string): Map where keys are keywords to find and values are replacement text

### Use Cases:

-   Creating fully populated documents from templates
-   Mail merge-style operations for complete document generation
-   Form filling where multiple fields need to be replaced
-   Template processing where all variables should be replaced at once
-   Contract generation with multiple client details

### Example Scenario:

```go
contractData := map[string]string{
    "CLIENT_NAME":        "John Smith",
    "CLIENT_ADDRESS":     "123 Main St, City, State 12345",
    "CLIENT_EMAIL":       "john.smith@email.com",
    "CONTRACT_DATE":      "2024-07-16",
    "CONTRACT_AMOUNT":    "$5,000",
    "PROJECT_NAME":       "Website Development",
    "PROJECT_DEADLINE":   "2024-08-30",
    "PAYMENT_TERMS":      "Net 30",
}
docx.ProcessDocxMulti("contract_template.docx", "john_smith_contract.docx", contractData)
```

## Function 3: ProcessDocxMultipleRecords

Creates multiple outputs (docx files) for each record, with keyword-replacement pairs applied simultaneously to each record

```go
records := []map[string]string{
		{
			"CLIENT_NAME":     "John Smith",
			"COMPANY_NAME":    "Smith Industries",
			"CONTRACT_DATE":   "2024-07-16",
			"CONTRACT_AMOUNT": "$5,000",
			"PROJECT_NAME":    "Website Development",
			"DEADLINE":        "2024-08-30",
		},
		{
			"CLIENT_NAME":     "John Doe",
			"COMPANY_NAME":    "Doe Industries",
			"CONTRACT_DATE":   "2024-03-06",
			"CONTRACT_AMOUNT": "$4,300",
			"PROJECT_NAME":    "Backend Development",
			"DEADLINE":        "2024-02-28",
		},
	}

	// Creates: contract_1.docx, contract_2.docx
	err := docx.ProcessDocxMultipleRecords("contract_template.docx", "./contracts", records, "contract_%d.docx")
```

### Parameters:

-   `inputPath` (string): Path to the source DOCX template file
-   `outputPath` (string): Path where the final DOCX file with all replacements will be saved
-   `record` ([]map[string]string): Dynamically sliced map where keys are keywords to find and values are replacement text
-   `fileNamePattern` (string): Printf-style pattern for naming files (e.g., `"contract_%d.docx`)

### Use Cases:

-   Creating multiple fully populated documents from a single template
-   Document generating for multiple records simultaneously

### Example Scenario:

```go
multipleInvoiceData := []map[string]string{
    {
        "INVOICE_DATE":       "2024-07-16",
        "CLIENT_NAME":        "ABC Corporation",
        "CLIENT_ADDRESS":     "456 Business Ave, Suite 100",
        "TOTAL_AMOUNT":       "$2,500.00",
        "DUE_DATE":           "2024-08-16",
        "DESCRIPTION":        "Web Development Services",
    },
    {
        "INVOICE_DATE":       "2024-07-23",
        "CLIENT_NAME":        "DEF Corporation",
        "CLIENT_ADDRESS":     "456 Business Ave, Suite 112",
        "TOTAL_AMOUNT":       "$1,230.00",
        "DUE_DATE":           "2024-09-01",
        "DESCRIPTION":        "Web Development Services",
    }
}

// Creates: invoice_1.docx, invoice_2.docx
docx.ProcessDocxMultipleRecords("invoice_template.docx", "./invoices", multipleInvoiceData, "invoice_%d.docx")
```

## Function 4: ProcessDocxMultipleRecordsWithNames

Creates multiple outputs (docx files) for each record, with keyword-replacement pairs applied simultaneously to each record, with a custom naming function

```go
records := []map[string]string{
    {
        "CLIENT_NAME":     "John Smith",
        "COMPANY_NAME":    "Smith Industries",
        "CONTRACT_DATE":   "2024-07-16",
        "CONTRACT_AMOUNT": "$5,000",
        "PROJECT_NAME":    "Website Development",
        "DEADLINE":        "2024-08-30",
    },
    {
        "CLIENT_NAME":     "John Doe",
        "COMPANY_NAME":    "Doe Industries",
        "CONTRACT_DATE":   "2024-03-06",
        "CONTRACT_AMOUNT": "$4,300",
        "PROJECT_NAME":    "Backend Development",
        "DEADLINE":        "2024-02-28",
    },
}

nameFunc := func(record map[string]string, index int) string {
    clientName := strings.ReplaceAll(record["CLIENT_NAME"], " ", "_")
    return fmt.Sprintf("contract_%s_%d.docx", strings.ToLower(clientName), index)
}

// Creates: contract_john_smith_1.docx, contract_john_doe_2.docx
err := docx.ProcessDocxMultipleRecordsWithNames("contract_template.docx", "./contracts", records, nameFunc)
```

### Parameters:

-   `inputPath` (string): Path to the source DOCX template file
-   `outputPath` (string): Path where the final DOCX file with all replacements will be saved
-   `record` ([]map[string]string): Dynamically sliced map where keys are keywords to find and values are replacement text
-   `nameFunc` func(map[string]string, int): Custom function to customize the naming of the output files

### Use Cases:

-   Creating multiple fully populated documents from a single template
-   Document generating for multiple records simultaneously
-   The use case calls for a custom naming convention instead of the simple Printf-style naming provided by ProcessDocxMultipleRecords

### Example Scenario:

```go
multipleInvoiceData := []map[string]string{
    {
        "INVOICE_NUMBER":     "INV-2024-001",
        "INVOICE_DATE":       "2024-07-16",
        "CLIENT_NAME":        "ABC Corporation",
        "CLIENT_ADDRESS":     "456 Business Ave, Suite 100",
        "TOTAL_AMOUNT":       "$2,500.00",
        "DUE_DATE":           "2024-08-16",
        "DESCRIPTION":        "Web Development Services",
    },
    {
        "INVOICE_NUMBER":     "INV-2024-002",
        "INVOICE_DATE":       "2024-07-23",
        "CLIENT_NAME":        "DEF Corporation",
        "CLIENT_ADDRESS":     "456 Business Ave, Suite 112",
        "TOTAL_AMOUNT":       "$1,230.00",
        "DUE_DATE":           "2024-09-01",
        "DESCRIPTION":        "Web Development Services",
    }
}

invoiceNamingFunction := func(map[string]string, index int) string {
    return record["INVOICE_NUMBER"] + ".docx"
}

// Creates: INV-2024-001.docx, INV-2024-002.docx
docx.ProcessDocxMultipleRecordsWithNames("invoice_template.docx", "./invoices", multipleInvoiceData, invoiceNamingFunction)
```

## Advanced Usage

### Error Handling:

```go
err := docx.ProcessDocxSingle("input.docx", "output.docx", "KEYWORD", "replacement")
if err != nil {
    log.Printf("Processing failed: %v", err)
    // Handle error appropriately
}

err = docx.ProcessDocxMulti("template.docx", "output.docx", replacements)
if err != nil {
    log.Printf("Batch processing failed: %v", err)
    // Handle error appropriately
}
```

### Checking File Existence:

```go
if _, err := os.Stat("template.docx"); os.IsNotExist(err) {
    log.Fatal("Template file does not exist")
}
```

### Creating Output Directories:

```go
outputPath := "./generated_documents/contract.docx"
outputDir := filepath.Dir(outputPath)
if err := os.MkdirAll(outputDir, 0755); err != nil {
    log.Printf("Failed to create output directory: %v", err)
}
```

## Best Practices

### Keyword Formatting:

-   Use consistent keyword formatting (e.g., `{{KEYWORD}}`, `KEYWORD`, `[KEYWORD]`)
-   Choose keywords that won't accidentally match regular text
-   Consider using unique delimiters to avoid false matches

### File Organization:

```go
// Organize by date
outputPath := fmt.Sprintf("./documents/%s/contract.docx", time.Now().Format("2006-01-02"))

// Organize by client
outputPath := fmt.Sprintf("./clients/%s/contract.docx", clientName)
```

### Template Design:

-   Keep templates simple and well-formatted
-   Use clear, descriptive keyword names
-   Test templates with sample data before production use
-   Consider using placeholder text that's obviously a placeholder and is not meant to be present in the final document (e.g., "REPLACE_WITH_CLIENT_NAME")

## Integration Examples

## CLI Tool Integration

```go
package main

import (
    "flag"
    "log"
    "github.com/siliconcatalyst/officeforge/docx"
)

func main() {
    var (
        input       = flag.String("input", "", "Input DOCX file")
        output      = flag.String("output", "", "Output DOCX file")
        keyword     = flag.String("keyword", "", "Keyword to replace")
        replacement = flag.String("replacement", "", "Replacement text")
    )
    flag.Parse()

    err := docx.ProcessDocxSingle(*input, *output, *keyword, *replacement)
    if err != nil {
        log.Fatal(err)
    }
}
```

### Usage Examples

```bash
# Single replacement
./cli-tool -input template.docx -output output.docx -keyword "{{NAME}}" -replacement "John Doe"

# Multiple replacements (build separate tool)
./multi-tool -input template.docx -output output.docx -config replacements.json

# Batch processing (build separate tool)
./batch-tool -input template.docx -output-dir ./outputs -records data.json
```

### Multi-Replacement CLI Example

```go
package main

import (
    "encoding/json"
    "flag"
    "os"
    "github.com/siliconcatalyst/officeforge/docx"

)

func main() {
    var (
        input  = flag.String("input", "", "Input DOCX file")
        output = flag.String("output", "", "Output DOCX file")
        config = flag.String("config", "", "JSON config file")
    )
    flag.Parse()

    // Read replacements from JSON
    data, _ := os.ReadFile(*config)
    var replacements map[string]string
    json.Unmarshal(data, &replacements)

    err := docx.ProcessDocxMulti(*input, *output, replacements)
    if err != nil {
        log.Fatal(err)
    }
}
```

## Configuration Examples

### Single Replacement Config

```json
{
	"template": "invoice_template.docx",
	"output": "invoice_001.docx",
	"keyword": "{{INVOICE_ID}}",
	"replacement": "INV-2025-001"
}
```

### Multi Replacement Config

```json
{
	"template": "contract_template.docx",
	"output": "contract_acme.docx",
	"replacements": {
		"{{CLIENT_NAME}}": "Acme Corporation",
		"{{DATE}}": "2025-07-16",
		"{{AMOUNT}}": "$50,000",
		"{{DURATION}}": "12 months",
		"{{PROJECT}}": "Website Development"
	}
}
```

### Batch Processing Config

```json
{
	"template": "certificate_template.docx",
	"output_dir": "./certificates",
	"file_name_pattern": "certificate_%s.docx",
	"records": [
		{
			"{{NAME}}": "John Doe",
			"{{EMAIL}}": "john@example.com",
			"{{COURSE}}": "Advanced Go Programming",
			"{{DATE}}": "2025-07-16"
		},
		{
			"{{NAME}}": "Jane Smith",
			"{{EMAIL}}": "jane@example.com",
			"{{COURSE}}": "Advanced Go Programming",
			"{{DATE}}": "2025-07-16"
		}
	]
}
```

## Error Handling Best Practices

### CLI Error Handling

```bash
# Robust error handling in bash
process_with_retry() {
    local max_attempts=3
    local attempt=1

    while [ $attempt -le $max_attempts ]; do
        echo "Attempt $attempt of $max_attempts"

        if ./docx-processor single --input "$1" --output "$2" --keyword "$3" --replacement "$4"; then
            echo "Success on attempt $attempt"
            return 0
        else
            echo "Attempt $attempt failed"
            ((attempt++))
            sleep 2
        fi
    done

    echo "All attempts failed"
    return 1
}
```

## Common Issues and Solutions

### Issue: Keywords not found

**Solution**: Verify keyword spelling and formatting in the template document

### Issue: Output files not created

**Solution**: Check that the output directory exists and has write permissions

### Issue: Malformed DOCX output

**Solution**: Ensure the input file is a valid DOCX file and not corrupted

### Issue: Partial replacements

**Solution**: Check for keyword conflicts or overlapping text patterns

### Issue: Memory usage with large files

**Solution**: Process documents individually rather than in large batches
