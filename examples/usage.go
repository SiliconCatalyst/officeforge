package examples

import (
	"encoding/json"
	"fmt"
	"log"
	"strings"

	"github.com/siliconcatalyst/officeforge/docx"
)

// Example 1: Single replacement
func ExampleUsage1() {

	invoiceNumber := 10223
	invoiceTemplate := "invoice_template.docx"
	invoiceOutput := fmt.Sprintf("invoice_no_%d.docx", invoiceNumber) // Output file name: invoice_no_10223.docx
	keyword := "CLIENT_NAME"
	replacement := "John Doe"

	err := docx.ProcessDocxSingle(invoiceTemplate, invoiceOutput, keyword, replacement)
	if err != nil {
		log.Printf("Error in single replacement: %v", err)
	}
}

// Example 2: Batch replacement (all replacements in one file)
func ExampleUsage2() {

	contractNumber := 13
	contractTemplate := "contract_template.docx"
	contractOutput := fmt.Sprintf("contract_no_%d.docx", contractNumber) // Output file name: contract_no_13.docx
	batchReplacements := map[string]string{
		"CLIENT_NAME":     "John Smith",
		"COMPANY_NAME":    "Smith Industries",
		"CONTRACT_DATE":   "2024-07-16",
		"CONTRACT_NUMBER": fmt.Sprint(contractNumber), // converts contractNumber (type int) to string
		"CONTRACT_AMOUNT": "$5,000",
		"PROJECT_NAME":    "Website Development",
		"DEADLINE":        "2024-08-30",
	}

	err := docx.ProcessDocxMulti(contractTemplate, contractOutput, batchReplacements)
	if err != nil {
		log.Printf("Error in batch replacement: %v", err)
	}
}

// Example 3: Batch replacement for each record
func ExampleUsage3() {

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
	if err != nil {
		log.Printf("Error in batch replacement: %v", err)
	}
}

// Example 4: Batch replacement for each record with a custom naming function
func ExampleUsage4() {

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
	if err != nil {
		log.Printf("Error in batch replacement: %v", err)
	}
}

// JSON Integration
func JsonExample() {
	jsonData := `[
    {
        "CLIENT_NAME": "John Smith",
        "COMPANY_NAME": "Smith Industries",
        "CONTRACT_DATE": "2024-07-16",
        "CONTRACT_AMOUNT": "$5,000",
        "PROJECT_NAME": "Website Development",
        "DEADLINE": "2024-08-30"
    },
    {
        "CLIENT_NAME": "John Doe",
        "COMPANY_NAME": "Doe Industries",
        "CONTRACT_DATE": "2024-03-06",
        "CONTRACT_AMOUNT": "$4,300",
        "PROJECT_NAME": "Website Development",
        "DEADLINE": "2024-02-28"
    }
]`

	var records []map[string]string
	err := json.Unmarshal([]byte(jsonData), &records)
	if err != nil {
		log.Print("Error in parsing JSON data")
	}

	err = docx.ProcessDocxMultipleRecords("template.docx", "./output", records, "document_%d.docx")
	if err != nil {
		log.Printf("Error in batch replacement with JSON data: %v", err)
	}
}
