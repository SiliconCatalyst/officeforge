# OfficeForge Examples

This directory contains working examples demonstrating all features of OfficeForge.

## 📁 Files Overview

```
examples/
├── example_usage.go          # Complete Go example using all 4 DOCX functions
├── sample_data.json          # Sample JSON data for batch processing
├── sample_data.csv           # Sample CSV data for batch processing
├── contract_template.docx    # Contract template with placeholders
├── invoice_template.docx     # Invoice template with placeholders
└── README.md                 # This file
```

## 🚀 Quick Start

### Prerequisites

**Place your template files** in this directory:
   - `contract_template.docx` - with placeholders like `{{CLIENT_NAME}}`, `{{CONTRACT_DATE}}`, etc.
   - `invoice_template.docx` - with placeholders like `{{INVOICE_NUMBER}}`, `{{CLIENT_NAME}}`, etc.

### Run the Examples

```bash
# From the examples directory
go run example_usage.go
```

This will demonstrate all 4 functions and create documents in the `output/` folder.

## 📚 Template Placeholders

### Contract Template Placeholders

Your `contract_template.docx` should include these placeholders:

```
{{CLIENT_NAME}}         - Client's full name
{{CLIENT_EMAIL}}        - Client's email address
{{CLIENT_PHONE}}        - Client's phone number
{{CLIENT_COMPANY}}      - Client's company name
{{CONTRACT_DATE}}       - Date of the contract
{{CONTRACT_AMOUNT}}     - Total contract value
{{PROJECT_NAME}}        - Name of the project
{{PROJECT_DEADLINE}}    - Project completion deadline
{{PAYMENT_TERMS}}       - Payment terms and conditions
```

**Example Contract Content:**
```
SERVICE AGREEMENT

This Service Agreement is entered into on {{CONTRACT_DATE}} between:

Service Provider: Your Company Name
Client: {{CLIENT_NAME}}
Company: {{CLIENT_COMPANY}}
Email: {{CLIENT_EMAIL}}
Phone: {{CLIENT_PHONE}}

PROJECT DETAILS
Project Name: {{PROJECT_NAME}}
Total Amount: {{CONTRACT_AMOUNT}}
Deadline: {{PROJECT_DEADLINE}}
Payment Terms: {{PAYMENT_TERMS}}
```

### Invoice Template Placeholders

Your `invoice_template.docx` should include these placeholders:

```
{{INVOICE_NUMBER}}      - Unique invoice identifier
{{INVOICE_DATE}}        - Date invoice was issued
{{CLIENT_NAME}}         - Client's name
{{CLIENT_ADDRESS}}      - Client's full address
{{CLIENT_EMAIL}}        - Client's email
{{ITEM_DESCRIPTION}}    - Description of service/product
{{QUANTITY}}            - Number of items/hours
{{UNIT_PRICE}}          - Price per unit
{{TOTAL_AMOUNT}}        - Total invoice amount
{{DUE_DATE}}            - Payment due date
{{PAYMENT_TERMS}}       - Payment terms
{{NOTES}}               - Additional notes
```

**Example Invoice Content:**
```
INVOICE

Invoice Number: {{INVOICE_NUMBER}}
Date: {{INVOICE_DATE}}
Due Date: {{DUE_DATE}}

BILL TO:
{{CLIENT_NAME}}
{{CLIENT_ADDRESS}}
Email: {{CLIENT_EMAIL}}

ITEMS:
Description: {{ITEM_DESCRIPTION}}
Quantity: {{QUANTITY}}
Unit Price: {{UNIT_PRICE}}

TOTAL AMOUNT: {{TOTAL_AMOUNT}}

Payment Terms: {{PAYMENT_TERMS}}
Notes: {{NOTES}}
```

## 🎯 Example Breakdown

### Example 1: Single Replacement

**Function:** `ProcessDocxSingle`

**Use Case:** Quick personalization

```go
docx.ProcessDocxSingle(
    "contract_template.docx",
    "output/contract_john_doe.docx",
    "{{CLIENT_NAME}}",
    "John Doe",
)
```

**Output:** One contract with `{{CLIENT_NAME}}` replaced with "John Doe"

---

### Example 2: Multiple Replacements

**Function:** `ProcessDocxMulti`

**Use Case:** Complete form filling

```go
invoiceData := map[string]string{
    "{{INVOICE_NUMBER}}": "INV-2024-001",
    "{{CLIENT_NAME}}":    "Acme Corporation",
    "{{TOTAL_AMOUNT}}":   "$5,000.00",
    // ... more fields
}

docx.ProcessDocxMulti(
    "invoice_template.docx",
    "output/invoice_acme_corp.docx",
    invoiceData,
)
```

**Output:** One complete invoice with all fields filled

---

### Example 3: Batch Processing

**Function:** `ProcessDocxMultipleRecords`

**Use Case:** Generate many documents with sequential naming

```go
records := []map[string]string{
    {"{{CLIENT_NAME}}": "Alice", "{{AMOUNT}}": "$10,000"},
    {"{{CLIENT_NAME}}": "Bob", "{{AMOUNT}}": "$15,000"},
    {"{{CLIENT_NAME}}": "Carol", "{{AMOUNT}}": "$25,000"},
}

docx.ProcessDocxMultipleRecords(
    "contract_template.docx",
    "output/batch_contracts",
    records,
    "contract_%d.docx",
)
```

**Output:** 
- `output/batch_contracts/contract_1.docx`
- `output/batch_contracts/contract_2.docx`
- `output/batch_contracts/contract_3.docx`

---

### Example 4: Custom Naming

**Function:** `ProcessDocxMultipleRecordsWithNames`

**Use Case:** Complex file naming logic

```go
nameFunc := func(record map[string]string, index int) string {
    invoiceNum := record["{{INVOICE_NUMBER}}"]
    clientName := strings.ReplaceAll(record["{{CLIENT_NAME}}"], " ", "_")
    return fmt.Sprintf("%s_%s.docx", invoiceNum, clientName)
}

docx.ProcessDocxMultipleRecordsWithNames(
    "invoice_template.docx",
    "output/custom_invoices",
    records,
    nameFunc,
)
```

**Output:**
- `output/custom_invoices/INV-2024-101_Global_Tech_Inc.docx`
- `output/custom_invoices/INV-2024-102_Startup_Ventures_LLC.docx`
- `output/custom_invoices/INV-2024-103_Enterprise_Holdings.docx`

## 🔧 Using Sample Data Files

### With JSON (sample_data.json)

```bash
# CLI
officeforge batch \
  --input contract_template.docx \
  --output ./output \
  --data sample_data.json

# Go
// Read and unmarshal sample_data.json
data, _ := os.ReadFile("sample_data.json")
var records []map[string]string
json.Unmarshal(data, &records)

docx.ProcessDocxMultipleRecords(
    "contract_template.docx",
    "./output",
    records,
    "contract_%d.docx",
)
```

### With CSV (sample_data.csv)

```bash
# CLI
officeforge batch \
  --input contract_template.docx \
  --output ./output \
  --data sample_data.csv \
  --pattern "{CLIENT_NAME}_contract.docx"
```

## 💡 Creating Your Own Templates

1. **Open Word** and create a new document
2. **Add your content** with placeholders
3. **Use consistent formatting**: `{{KEYWORD}}` or `[[KEYWORD]]`
4. **Save as .docx**
5. **Test** with sample data

**Tips:**
- Use ALL CAPS for placeholders (easier to spot)
- Keep keywords descriptive: `{{CLIENT_NAME}}` not `{{N}}`
- Test with the longest expected values to check formatting
- Use tables for structured data

## 🎨 Customization Ideas

### Different Document Types

- **Certificates**: `{{RECIPIENT_NAME}}`, `{{COURSE_NAME}}`, `{{DATE}}`
- **Employment Letters**: `{{EMPLOYEE_NAME}}`, `{{POSITION}}`, `{{START_DATE}}`
- **Proposals**: `{{PROJECT_NAME}}`, `{{BUDGET}}`, `{{TIMELINE}}`
- **Reports**: `{{REPORT_DATE}}`, `{{METRICS}}`, `{{SUMMARY}}`

### Advanced Patterns

```go
// Date-based naming
func(record map[string]string, index int) string {
    date := time.Now().Format("2006-01-02")
    return fmt.Sprintf("%s_report_%d.docx", date, index)
}

// Client-based organization
func(record map[string]string, index int) string {
    client := record["{{CLIENT_NAME}}"]
    return fmt.Sprintf("clients/%s/contract.docx", client)
}

// Invoice numbering
func(record map[string]string, index int) string {
    return fmt.Sprintf("INV-%04d.docx", index+1)
}
```

## 🐛 Troubleshooting

### Template not found
```
Error: open contract_template.docx: no such file or directory
```
**Solution:** Ensure template files are in the examples directory

### Output directory error
```
Error: failed to create output file
```
**Solution:** Create output directories first:
```bash
mkdir -p output/batch_contracts output/custom_invoices
```

### Placeholders not replaced
**Solution:** 
- Check exact spelling in template
- Ensure no extra spaces: `{{ NAME }}` vs `{{NAME}}`
- Verify placeholder format matches your data

### Formatting issues
**Solution:**
- Keep placeholders on single lines
- Don't split placeholders across formatting changes
- Use "Keep text together" in Word for important sections

## 📖 Further Reading

- [Main Documentation](../README.md)
- [API Reference](https://pkg.go.dev/github.com/siliconcatalyst/officeforge)
- [CLI Usage Guide](../README.md#cli-documentation)

## 🤝 Contributing Examples

Have a cool use case? Submit a PR with:
1. New template file
2. Sample data
3. Working Go example
4. Documentation

---

**Questions?** Open an issue at https://github.com/siliconcatalyst/officeforge/issues