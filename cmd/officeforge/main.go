package main

import (
	"fmt"
	"os"
)

func main() {
	if len(os.Args) < 2 {
		printUsage()
		os.Exit(1)
	}

	command := os.Args[1]

	switch command {
	// DOCX commands
	case "docx-single":
		handleDocxSingle(os.Args[2:])
	case "docx-multi":
		handleDocxMulti(os.Args[2:])
	case "docx-batch":
		handleDocxBatch(os.Args[2:])
	case "docx-check":
		handleDocxCheck(os.Args[2:])

	// XLSX commands
	case "xlsx-single":
		handleXlsxSingle(os.Args[2:])
	case "xlsx-multi":
		handleXlsxMulti(os.Args[2:])
	case "xlsx-batch":
		handleXlsxBatch(os.Args[2:])
	case "xlsx-check":
		handleXlsxCheck(os.Args[2:])

	// PPTX commands
	case "pptx-single":
		handlePptxSingle(os.Args[2:])
	case "pptx-multi":
		handlePptxMulti(os.Args[2:])
	case "pptx-batch":
		handlePptxBatch(os.Args[2:])
	case "pptx-check":
		handlePptxCheck(os.Args[2:])

	// Other commands
	case "version":
		printVersion()
	case "help", "-h", "--help":
		printUsage()
	default:
		fmt.Printf("Unknown command: %s\n\n", command)
		printUsage()
		os.Exit(1)
	}
}

func printUsage() {
	fmt.Println(`OfficeForge - Generate Office documents from the command line

Usage:
  officeforge <command> [options]

Commands:
  DOCX (Word Documents):
    docx-single      Replace a single keyword in a template
    docx-multi       Replace multiple keywords in a template
    docx-batch       Generate multiple documents from a template
    docx-check       Check if keywords exist in a document

  XLSX (Excel Spreadsheets):
    xlsx-single      Replace a single keyword in a template
    xlsx-multi       Replace multiple keywords in a template
    xlsx-batch       Generate multiple spreadsheets from a template
    xlsx-check       Check if keywords exist in a spreadsheet

  PPTX (PowerPoint Presentations):
    pptx-single      Replace a single keyword in a template
    pptx-multi       Replace multiple keywords in a template
    pptx-batch       Generate multiple presentations from a template
    pptx-check       Check if keywords exist in a presentation

  Other:
    version          Show version
    help             Show this help message

Examples:
  # Replace single keyword in Word document
  officeforge docx-single --input template.docx --output result.docx --key "{{NAME}}" --value "John Doe"

  # Replace multiple keywords in Excel spreadsheet
  officeforge xlsx-multi --input template.xlsx --output result.xlsx --data replacements.json

  # Generate multiple PowerPoint presentations from CSV
  officeforge pptx-batch --input template.pptx --output ./output --data records.csv

  # Generate documents with custom naming pattern
  officeforge docx-batch --input template.docx --output ./output --data records.csv --pattern "{name}_{id}.docx"

  # Check for specific keywords
  officeforge docx-check --input doc.docx --keys "TOTAL_COST,DATE,SIGNATURE"

For more information, visit: https://github.com/siliconcatalyst/officeforge`)
}
