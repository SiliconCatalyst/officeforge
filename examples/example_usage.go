package main

import (
	"fmt"
	"log"
	"strings"

	"github.com/siliconcatalyst/officeforge/docx"
)

func main() {
	fmt.Println("🚀 OfficeForge Examples - Demonstrating all 4 functions")

	// Example 1: ProcessDocxSingle - Single Contract
	example1()

	// Example 2: ProcessDocxMulti - Complete Invoice
	example2()

	// Example 3: ProcessDocxMultipleRecords - Batch Contracts
	example3()

	// Example 4: ProcessDocxMultipleRecordsWithNames - Custom Named Invoices
	example4()

	fmt.Println("\n✅ All examples completed successfully!")
	fmt.Println("📁 Check the 'output' folder for generated documents")
}

// Example 1: Single keyword replacement
// Use Case: Quick personalization of a single document
func example1() {
	fmt.Println("1️⃣  Example 1: Single Replacement (ProcessDocxSingle)")
	fmt.Println("   Creating a personalized contract with one keyword...")

	err := docx.ProcessDocxSingle(
		"contract_template.docx",
		"output/contract_john_doe.docx",
		"{{CLIENT_NAME}}",
		"John Doe",
	)

	if err != nil {
		log.Printf("   ❌ Error: %v\n", err)
		return
	}

	fmt.Println("   ✓ Created: output/contract_john_doe.docx")
	fmt.Println("   ✓ Replaced: {{CLIENT_NAME}} → John Doe")
}

// Example 2: Multiple keyword replacements
// Use Case: Fill out a complete form/document with multiple fields
func example2() {
	fmt.Println("2️⃣  Example 2: Multiple Replacements (ProcessDocxMulti)")
	fmt.Println("   Creating a complete invoice with all details...")

	invoiceData := map[string]string{
		"{{INVOICE_NUMBER}}":   "INV-2024-001",
		"{{INVOICE_DATE}}":     "December 04, 2024",
		"{{CLIENT_NAME}}":      "Acme Corporation",
		"{{CLIENT_ADDRESS}}":   "123 Business Ave, Suite 100, Tech City, TC 12345",
		"{{CLIENT_EMAIL}}":     "billing@acmecorp.com",
		"{{ITEM_DESCRIPTION}}": "Professional Services - Q4 2024",
		"{{QUANTITY}}":         "1",
		"{{UNIT_PRICE}}":       "$5,000.00",
		"{{TOTAL_AMOUNT}}":     "$5,000.00",
		"{{DUE_DATE}}":         "January 03, 2025",
		"{{PAYMENT_TERMS}}":    "Net 30 days",
		"{{NOTES}}":            "Thank you for your business!",
	}

	err := docx.ProcessDocxMulti(
		"invoice_template.docx",
		"output/invoice_acme_corp.docx",
		invoiceData,
	)

	if err != nil {
		log.Printf("   ❌ Error: %v\n", err)
		return
	}

	fmt.Println("   ✓ Created: output/invoice_acme_corp.docx")
	fmt.Printf("   ✓ Replaced: %d keywords\n\n", len(invoiceData))
}

// Example 3: Batch processing with pattern naming
// Use Case: Generate multiple contracts from a list of clients
func example3() {
	fmt.Println("3️⃣  Example 3: Batch Processing (ProcessDocxMultipleRecords)")
	fmt.Println("   Generating contracts for multiple clients...")

	// This data could come from a CSV, database, or API
	contractRecords := []map[string]string{
		{
			"{{CLIENT_NAME}}":      "Alice Johnson",
			"{{CLIENT_EMAIL}}":     "alice@techstartup.com",
			"{{CLIENT_PHONE}}":     "555-0101",
			"{{CLIENT_COMPANY}}":   "Tech Startup Inc.",
			"{{CONTRACT_DATE}}":    "December 04, 2024",
			"{{CONTRACT_AMOUNT}}":  "$10,000",
			"{{PROJECT_NAME}}":     "Website Redesign",
			"{{PROJECT_DEADLINE}}": "January 31, 2025",
			"{{PAYMENT_TERMS}}":    "50% upfront, 50% on completion",
		},
		{
			"{{CLIENT_NAME}}":      "Bob Williams",
			"{{CLIENT_EMAIL}}":     "bob@consulting.com",
			"{{CLIENT_PHONE}}":     "555-0102",
			"{{CLIENT_COMPANY}}":   "Williams Consulting LLC",
			"{{CONTRACT_DATE}}":    "December 04, 2024",
			"{{CONTRACT_AMOUNT}}":  "$15,000",
			"{{PROJECT_NAME}}":     "Mobile App Development",
			"{{PROJECT_DEADLINE}}": "February 28, 2025",
			"{{PAYMENT_TERMS}}":    "Monthly installments",
		},
		{
			"{{CLIENT_NAME}}":      "Carol Davis",
			"{{CLIENT_EMAIL}}":     "carol@enterprise.com",
			"{{CLIENT_PHONE}}":     "555-0103",
			"{{CLIENT_COMPANY}}":   "Enterprise Solutions Corp.",
			"{{CONTRACT_DATE}}":    "December 04, 2024",
			"{{CONTRACT_AMOUNT}}":  "$25,000",
			"{{PROJECT_NAME}}":     "Database Migration",
			"{{PROJECT_DEADLINE}}": "March 15, 2025",
			"{{PAYMENT_TERMS}}":    "Milestone-based payments",
		},
	}

	// Pattern: contract_1.docx, contract_2.docx, contract_3.docx
	err := docx.ProcessDocxMultipleRecords(
		"contract_template.docx",
		"output/batch_contracts",
		contractRecords,
		"contract_%d.docx",
	)

	if err != nil {
		log.Printf("   ❌ Error: %v\n", err)
		return
	}

	fmt.Printf("   ✓ Created: %d contracts in output/batch_contracts/\n", len(contractRecords))
	fmt.Println("   ✓ Files: contract_1.docx, contract_2.docx, contract_3.docx")
}

// Example 4: Batch processing with custom naming function
// Use Case: Generate invoices with custom file names based on invoice numbers
func example4() {
	fmt.Println("4️⃣  Example 4: Custom Naming (ProcessDocxMultipleRecordsWithNames)")
	fmt.Println("   Generating invoices with custom file names...")

	invoiceRecords := []map[string]string{
		{
			"{{INVOICE_NUMBER}}":   "INV-2024-101",
			"{{INVOICE_DATE}}":     "December 01, 2024",
			"{{CLIENT_NAME}}":      "Global Tech Inc.",
			"{{CLIENT_ADDRESS}}":   "456 Innovation Drive, Silicon Valley, CA 94025",
			"{{CLIENT_EMAIL}}":     "accounts@globaltech.com",
			"{{ITEM_DESCRIPTION}}": "Cloud Infrastructure Setup",
			"{{QUANTITY}}":         "1",
			"{{UNIT_PRICE}}":       "$8,500.00",
			"{{TOTAL_AMOUNT}}":     "$8,500.00",
			"{{DUE_DATE}}":         "December 31, 2024",
			"{{PAYMENT_TERMS}}":    "Net 30 days",
			"{{NOTES}}":            "Payment due upon receipt",
		},
		{
			"{{INVOICE_NUMBER}}":   "INV-2024-102",
			"{{INVOICE_DATE}}":     "December 02, 2024",
			"{{CLIENT_NAME}}":      "Startup Ventures LLC",
			"{{CLIENT_ADDRESS}}":   "789 Entrepreneur Way, Austin, TX 78701",
			"{{CLIENT_EMAIL}}":     "billing@startupventures.com",
			"{{ITEM_DESCRIPTION}}": "MVP Development Services",
			"{{QUANTITY}}":         "1",
			"{{UNIT_PRICE}}":       "$12,000.00",
			"{{TOTAL_AMOUNT}}":     "$12,000.00",
			"{{DUE_DATE}}":         "January 01, 2025",
			"{{PAYMENT_TERMS}}":    "Net 30 days",
			"{{NOTES}}":            "Includes 3 months support",
		},
		{
			"{{INVOICE_NUMBER}}":   "INV-2024-103",
			"{{INVOICE_DATE}}":     "December 03, 2024",
			"{{CLIENT_NAME}}":      "Enterprise Holdings",
			"{{CLIENT_ADDRESS}}":   "321 Corporate Blvd, New York, NY 10001",
			"{{CLIENT_EMAIL}}":     "ap@enterpriseholdings.com",
			"{{ITEM_DESCRIPTION}}": "System Integration & Training",
			"{{QUANTITY}}":         "1",
			"{{UNIT_PRICE}}":       "$18,500.00",
			"{{TOTAL_AMOUNT}}":     "$18,500.00",
			"{{DUE_DATE}}":         "January 02, 2025",
			"{{PAYMENT_TERMS}}":    "Net 45 days",
			"{{NOTES}}":            "PO #: PO-2024-9876",
		},
	}

	// Custom naming function: uses invoice number and client name
	nameFunc := func(record map[string]string, index int) string {
		invoiceNum := record["{{INVOICE_NUMBER}}"]
		clientName := record["{{CLIENT_NAME}}"]

		// Clean client name for filename (remove spaces and special chars)
		cleanName := strings.ReplaceAll(clientName, " ", "_")
		cleanName = strings.ReplaceAll(cleanName, ".", "")
		cleanName = strings.ReplaceAll(cleanName, ",", "")

		return fmt.Sprintf("%s_%s.docx", invoiceNum, cleanName)
	}

	err := docx.ProcessDocxMultipleRecordsWithNames(
		"invoice_template.docx",
		"output/custom_invoices",
		invoiceRecords,
		nameFunc,
	)

	if err != nil {
		log.Printf("   ❌ Error: %v\n", err)
		return
	}

	fmt.Printf("   ✓ Created: %d invoices in output/custom_invoices/", len(invoiceRecords))
	fmt.Println("   ✓ Files:")
	for i, record := range invoiceRecords {
		fileName := nameFunc(record, i)
		fmt.Printf("      - %s\n", fileName)
	}
}
