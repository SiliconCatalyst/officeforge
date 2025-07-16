package examples

import (
	"log"

	"github.com/siliconcatalyst/officeforge/docx"
)

func ExampleUsage() {
	// Example 1: Single replacement
	err := docx.ProcessDocxSingle("input.docx", "output_single.docx", "PLACEHOLDER", "John Doe")
	if err != nil {
		log.Printf("Error in single replacement: %v", err)
	}

	// Example 2: Batch replacement (all replacements in one file)
	batchReplacements := map[string]string{
		"CLIENT_NAME":     "John Smith",
		"COMPANY_NAME":    "Smith Industries",
		"CONTRACT_DATE":   "2024-07-16",
		"CONTRACT_AMOUNT": "$5,000",
		"PROJECT_NAME":    "Website Development",
		"DEADLINE":        "2024-08-30",
	}

	err = docx.ProcessDocxBatch("contract_template.docx", "completed_contract.docx", batchReplacements)
	if err != nil {
		log.Printf("Error in batch replacement: %v", err)
	}
}
