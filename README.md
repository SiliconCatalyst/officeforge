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

## Function 1: ProcessDocxSingle

Creates a single output DOCX file with one keyword-replacement pair applied.

```go
err := ProcessDocxSingle("contract.docx", "contract_john.docx", "CLIENT_NAME", "John Smith")
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
ProcessDocxSingle("template_contract.docx", "smith_contract.docx", "{{CLIENT}}", "John Smith")

// Create a customized letter
ProcessDocxSingle("letter_template.docx", "welcome_letter.docx", "COMPANY_NAME", "Tech Solutions Inc.")

// Replace dates in documents
ProcessDocxSingle("report.docx", "monthly_report.docx", "REPORT_DATE", "July 2024")
```

## Function 2: ProcessDocxBatch

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
err := ProcessDocxBatch("contract_template.docx", "completed_contract.docx", batchReplacements)
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

### Example Scenarios:

#### Complete Contract Generation:

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
ProcessDocxBatch("contract_template.docx", "john_smith_contract.docx", contractData)
```

#### Invoice Generation:

```go
invoiceData := map[string]string{
    "INVOICE_NUMBER":     "INV-2024-001",
    "INVOICE_DATE":       "2024-07-16",
    "CLIENT_NAME":        "ABC Corporation",
    "CLIENT_ADDRESS":     "456 Business Ave, Suite 100",
    "TOTAL_AMOUNT":       "$2,500.00",
    "DUE_DATE":           "2024-08-16",
    "DESCRIPTION":        "Web Development Services",
}
ProcessDocxBatch("invoice_template.docx", "invoice_001.docx", invoiceData)
```

#### Employee Document Generation:

```go
employeeData := map[string]string{
    "EMPLOYEE_NAME":      "Jane Doe",
    "EMPLOYEE_ID":        "EMP-2024-001",
    "DEPARTMENT":         "Engineering",
    "START_DATE":         "2024-07-16",
    "SALARY":             "$75,000",
    "MANAGER_NAME":       "Bob Johnson",
    "OFFICE_LOCATION":    "New York",
}
ProcessDocxBatch("employee_contract.docx", "jane_doe_contract.docx", employeeData)
```

## Advanced Usage

### Error Handling:

```go
err := ProcessDocxSingle("input.docx", "output.docx", "KEYWORD", "replacement")
if err != nil {
    log.Printf("Processing failed: %v", err)
    // Handle error appropriately
}

err = ProcessDocxBatch("template.docx", "output.docx", replacements)
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
-   Consider using placeholder text that's obviously fake (e.g., "REPLACE_WITH_CLIENT_NAME")

### Performance Considerations:

```go
// For large documents, consider processing in chunks
// Use ProcessDocxBatch for efficiency when replacing multiple keywords
func ProcessLargeDocument(templatePath, outputPath string, data map[string]string) error {
    // ProcessDocxBatch is more efficient than multiple ProcessDocxSingle calls
    return ProcessDocxBatch(templatePath, outputPath, data)
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

## Integration Examples

### Web Service Integration:

```go
func handleSingleReplacement(w http.ResponseWriter, r *http.Request) {
    keyword := r.FormValue("keyword")
    replacement := r.FormValue("replacement")

    outputPath := fmt.Sprintf("./temp/%s.docx", uuid.New().String())
    err := ProcessDocxSingle("template.docx", outputPath, keyword, replacement)

    if err != nil {
        http.Error(w, err.Error(), http.StatusInternalServerError)
        return
    }

    http.ServeFile(w, r, outputPath)
}

func handleBatchReplacement(w http.ResponseWriter, r *http.Request) {
    // Parse JSON request body containing replacements map
    var replacements map[string]string
    if err := json.NewDecoder(r.Body).Decode(&replacements); err != nil {
        http.Error(w, err.Error(), http.StatusBadRequest)
        return
    }

    outputPath := fmt.Sprintf("./temp/%s.docx", uuid.New().String())
    err := ProcessDocxBatch("template.docx", outputPath, replacements)

    if err != nil {
        http.Error(w, err.Error(), http.StatusInternalServerError)
        return
    }

    http.ServeFile(w, r, outputPath)
}
```

### CLI Tool Integration:

```go
func main() {
    if len(os.Args) < 2 {
        log.Fatal("Usage: docx-processor <command> [args...]")
    }

    command := os.Args[1]

    switch command {
    case "single":
        if len(os.Args) < 6 {
            log.Fatal("Usage: docx-processor single <input> <output> <keyword> <replacement>")
        }
        err := ProcessDocxSingle(os.Args[2], os.Args[3], os.Args[4], os.Args[5])
        if err != nil {
            log.Fatal(err)
        }
        fmt.Println("Single replacement completed successfully!")

    case "batch":
        if len(os.Args) < 5 {
            log.Fatal("Usage: docx-processor batch <input> <output> <replacements.json>")
        }

        // Read replacements from JSON file
        data, err := os.ReadFile(os.Args[4])
        if err != nil {
            log.Fatal(err)
        }

        var replacements map[string]string
        if err := json.Unmarshal(data, &replacements); err != nil {
            log.Fatal(err)
        }

        err = ProcessDocxBatch(os.Args[2], os.Args[3], replacements)
        if err != nil {
            log.Fatal(err)
        }
        fmt.Println("Batch replacement completed successfully!")

    default:
        log.Fatal("Unknown command. Use 'single' or 'batch'")
    }
}
```

### Configuration File Support:

```go
type DocumentConfig struct {
    Template     string            `json:"template"`
    Output       string            `json:"output"`
    Replacements map[string]string `json:"replacements"`
}

func ProcessFromConfig(configPath string) error {
    data, err := os.ReadFile(configPath)
    if err != nil {
        return err
    }

    var config DocumentConfig
    if err := json.Unmarshal(data, &config); err != nil {
        return err
    }

    return ProcessDocxBatch(config.Template, config.Output, config.Replacements)
}
```

Example configuration file (`config.json`):

```json
{
	"template": "contract_template.docx",
	"output": "final_contract.docx",
	"replacements": {
		"CLIENT_NAME": "John Smith",
		"CONTRACT_DATE": "2024-07-16",
		"AMOUNT": "$5,000"
	}
}
```
