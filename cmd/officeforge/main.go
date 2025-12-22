package main

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/siliconcatalyst/officeforge/docx"
	"github.com/siliconcatalyst/officeforge/internal"
)

var Version = "v0.0.0"

func main() {
	if len(os.Args) < 2 {
		printUsage()
		os.Exit(1)
	}

	command := os.Args[1]

	switch command {
	case "docx-single":
		handleDocxSingle(os.Args[2:])
	case "docx-multi":
		handleDocxMulti(os.Args[2:])
	case "docx-batch":
		handleDocxBatch(os.Args[2:])
	case "docx-check":
		handleDocxCheck(os.Args[2:])
	case "version":
		fmt.Println("OfficeForge", Version)
	case "help", "-h", "--help":
		printUsage()
	default:
		fmt.Printf("Unknown command: %s\n\n", command)
		printUsage()
		os.Exit(1)
	}
}

func printUsage() {
	fmt.Println(`OfficeForge - Generate Word documents from the command line

Usage:
  officeforge <command> [options]

Commands:
  docx-single      Replace a single keyword in a template
  docx-multi       Replace multiple keywords in a template
  docx-batch       Generate multiple documents from a template
  docx-check	   Check if keywords exist in a document
  version     Show version
  help        Show this help message

Examples:
  # Replace single keyword
  officeforge docx-single --input template.docx --output result.docx --key "{{NAME}}" --value "John Doe"

  # Replace multiple keywords from JSON
  officeforge docx-multi --input template.docx --output result.docx --data replacements.json

  # Generate multiple documents from CSV
  officeforge docx-batch --input template.docx --output ./output --data records.csv

  # Generate multiple documents with custom internal
  officeforge docx-batch --input template.docx --output ./output --data records.csv --pattern "{name}_{id}.docx"

  # Check for specific keywords
  officeforge docx-check --input doc.docx --keys "TOTAL_COST,DATE,SIGNATURE"
  
  # Check using a JSON file of keys
  officeforge docx-check --input doc.docx --data keys.json

For more information, visit: https://github.com/siliconcatalyst/officeforge`)
}

func handleDocxSingle(args []string) {
	if len(args) < 8 {
		fmt.Println("Error: Missing required arguments")
		fmt.Println("\nUsage:")
		fmt.Println("  officeforge single --input <template> --output <file> --key <keyword> --value <replacement>")
		os.Exit(1)
	}

	var inputPath, outputPath, keyword, replacement string

	for i := 0; i < len(args); i++ {
		switch args[i] {
		case "--input", "-i":
			if i+1 < len(args) {
				inputPath = args[i+1]
				i++
			}
		case "--output", "-o":
			if i+1 < len(args) {
				outputPath = args[i+1]
				i++
			}
		case "--key", "-k":
			if i+1 < len(args) {
				keyword = args[i+1]
				i++
			}
		case "--value", "-v":
			if i+1 < len(args) {
				replacement = args[i+1]
				i++
			}
		}
	}

	if inputPath == "" || outputPath == "" || keyword == "" || replacement == "" {
		fmt.Println("Error: All flags (--input, --output, --key, --value) are required")
		os.Exit(1)
	}

	err := docx.ProcessDocxSingle(inputPath, outputPath, keyword, replacement)
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("✓ Document created: %s\n", outputPath)
}

func handleDocxMulti(args []string) {
	if len(args) < 6 {
		fmt.Println("Error: Missing required arguments")
		fmt.Println("\nUsage:")
		fmt.Println("  officeforge multi --input <template> --output <file> --data <json_file>")
		os.Exit(1)
	}

	var inputPath, outputPath, dataPath string

	for i := 0; i < len(args); i++ {
		switch args[i] {
		case "--input", "-i":
			if i+1 < len(args) {
				inputPath = args[i+1]
				i++
			}
		case "--output", "-o":
			if i+1 < len(args) {
				outputPath = args[i+1]
				i++
			}
		case "--data", "-d":
			if i+1 < len(args) {
				dataPath = args[i+1]
				i++
			}
		}
	}

	if inputPath == "" || outputPath == "" || dataPath == "" {
		fmt.Println("Error: All flags (--input, --output, --data) are required")
		os.Exit(1)
	}

	// Read JSON file
	data, err := os.ReadFile(dataPath)
	if err != nil {
		fmt.Printf("Error reading data file: %v\n", err)
		os.Exit(1)
	}

	var replacements map[string]string
	err = json.Unmarshal(data, &replacements)
	if err != nil {
		fmt.Printf("Error parsing JSON: %v\n", err)
		os.Exit(1)
	}

	err = docx.ProcessDocxMulti(inputPath, outputPath, replacements)
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("✓ Document created: %s\n", outputPath)
	fmt.Printf("  Replaced %d keywords\n", len(replacements))
}

func handleDocxBatch(args []string) {
	if len(args) < 6 {
		fmt.Println("Error: Missing required arguments")
		fmt.Println("\nUsage:")
		fmt.Println("  officeforge docx-batch --input <template> --output <directory> --data <csv_or_json_file> [--pattern <pattern>]")
		fmt.Println("\nPattern examples:")
		fmt.Printf("  --pattern \"contract_%%d.docx\"           Sequential: contract_1.docx, contract_2.docx\n")
		fmt.Println("  --pattern \"{NAME}_contract.docx\"      From data: Alice_contract.docx, Bob_contract.docx")
		fmt.Println("  --pattern \"{ID}_{COMPANY}.docx\"       Multiple fields: 001_Acme.docx, 002_TechCorp.docx")
		fmt.Println("  --pattern \"{NAME}_{INDEX}.docx\"       Combine data and index: Alice_1.docx, Bob_2.docx")
		os.Exit(1)
	}

	var inputPath, outputDir, dataPath, pattern string

	for i := 0; i < len(args); i++ {
		switch args[i] {
		case "--input", "-i":
			if i+1 < len(args) {
				inputPath = args[i+1]
				i++
			}
		case "--output", "-o":
			if i+1 < len(args) {
				outputDir = args[i+1]
				i++
			}
		case "--data", "-d":
			if i+1 < len(args) {
				dataPath = args[i+1]
				i++
			}
		case "--pattern", "-p":
			if i+1 < len(args) {
				pattern = args[i+1]
				i++
			}
		}
	}

	if inputPath == "" || outputDir == "" || dataPath == "" {
		fmt.Println("Error: All flags (--input, --output, --data) are required")
		os.Exit(1)
	}

	// Create output directory if it doesn't exist
	err := os.MkdirAll(outputDir, 0755)
	if err != nil {
		fmt.Printf("Error creating output directory: %v\n", err)
		os.Exit(1)
	}

	// Determine file type and read data
	ext := strings.ToLower(filepath.Ext(dataPath))
	var records []map[string]string

	switch ext {
	case ".json":
		records, err = readJSONRecords(dataPath)
	case ".csv":
		records, err = readCSVRecords(dataPath)
	default:
		fmt.Printf("Error: Unsupported data file format: %s (use .json or .csv)\n", ext)
		os.Exit(1)
	}

	if err != nil {
		fmt.Printf("Error reading data file: %v\n", err)
		os.Exit(1)
	}

	if len(records) == 0 {
		fmt.Println("Error: No records found in data file")
		os.Exit(1)
	}

	// Validate pattern if provided
	if pattern != "" {
		if err := internal.ValidatePattern(pattern, records[0]); err != nil {
			fmt.Printf("Error: Invalid pattern - %v\n", err)
			os.Exit(1)
		}

		// Show what pattern type is being used
		patternType := internal.DetectPatternType(pattern)
		switch patternType {
		case internal.PatternTypeSequential:
			fmt.Printf("Using sequential pattern: %s\n", pattern)
		case internal.PatternTypeData:
			placeholders := internal.ExtractPlaceholders(pattern)
			fmt.Printf("Using data-based pattern with fields: %v\n", placeholders)
		}
	}

	// Process documents using the pattern
	err = docx.ProcessDocxMultipleRecords(inputPath, outputDir, records, pattern)
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("✓ Generated %d documents in: %s\n", len(records), outputDir)
}

func handleDocxCheck(args []string) {
	if len(args) < 4 {
		fmt.Println("Error: Missing required arguments")
		fmt.Println("\nUsage:")
		fmt.Println("  officeforge docx-check --input <file> [--keys \"K1,K2\" | --data <json_file>] --json")
		os.Exit(1)
	}

	var inputPath, keysString, dataPath string
	var outputJson bool

	for i := 0; i < len(args); i++ {
		switch args[i] {
		case "--input", "-i":
			if i+1 < len(args) {
				inputPath = args[i+1]
				i++
			}
		case "--keys", "-k":
			if i+1 < len(args) {
				keysString = args[i+1]
				i++
			}
		case "--data", "-d":
			if i+1 < len(args) {
				dataPath = args[i+1]
				i++
			}
		case "--json":
			outputJson = true
		}
	}

	if inputPath == "" || (keysString == "" && dataPath == "") {
		fmt.Println("Error: --input and either --keys or --data are required")
		os.Exit(1)
	}

	var keywords []string

	// Load keywords from JSON or Comma Separated String
	if dataPath != "" {
		data, err := os.ReadFile(dataPath)
		if err != nil {
			fmt.Printf("Error reading data file: %v\n", err)
			os.Exit(1)
		}
		// Try parsing as array first, then as map keys
		var keyList []string
		if err := json.Unmarshal(data, &keyList); err == nil {
			keywords = keyList
		} else {
			var keyMap map[string]interface{}
			if err := json.Unmarshal(data, &keyMap); err == nil {
				for k := range keyMap {
					keywords = append(keywords, k)
				}
			}
		}
	} else {
		keywords = strings.Split(keysString, ",")
		for i := range keywords {
			keywords[i] = strings.TrimSpace(keywords[i])
		}
	}

	// Call the logic function (which you'll add to your docx package)
	results, err := internal.ValidateKeywords(inputPath, keywords)
	if err != nil {
		// If there's a system error, we should still handle it
		if outputJson {
			fmt.Printf(`{"error": "%v"}`, err)
		} else {
			fmt.Printf("Error: %v\n", err)
		}
		os.Exit(1)
	}

	// Calculate foundCount for the summary
	foundCount := 0
	for _, found := range results {
		if found {
			foundCount++
		}
	}

	// --- OUTPUT LOGIC ---
	if outputJson {
		// ONLY print JSON so the UI can parse it easily
		jsonBytes, _ := json.Marshal(results)
		fmt.Println(string(jsonBytes))
	} else {
		// Human-readable table
		fmt.Printf("\nKeyword Check Results for: %s\n", inputPath)
		fmt.Println(strings.Repeat("-", 40))
		for _, k := range keywords {
			status := "✘ Missing"
			if results[k] {
				status = "✓ Found"
			}
			fmt.Printf("%-25s %s\n", k, status)
		}
		fmt.Println(strings.Repeat("-", 40))
		fmt.Printf("Summary: %d/%d keywords present\n", foundCount, len(keywords))
	}

	// Exit with 1 if any are missing
	if foundCount < len(keywords) {
		os.Exit(1)
	}
}

func readJSONRecords(path string) ([]map[string]string, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}

	var records []map[string]string
	err = json.Unmarshal(data, &records)
	if err != nil {
		return nil, err
	}

	return records, nil
}

func readCSVRecords(path string) ([]map[string]string, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	reader := csv.NewReader(file)
	rows, err := reader.ReadAll()
	if err != nil {
		return nil, err
	}

	if len(rows) < 2 {
		return nil, fmt.Errorf("CSV must have at least a header row and one data row")
	}

	headers := rows[0]
	var records []map[string]string

	for i := 1; i < len(rows); i++ {
		if len(rows[i]) != len(headers) {
			continue // Skip malformed rows
		}

		record := make(map[string]string)
		for j, header := range headers {
			record[header] = rows[i][j]
		}
		records = append(records, record)
	}

	return records, nil
}
