package main

import (
	"bytes"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/siliconcatalyst/officeforge/internal"
	"github.com/siliconcatalyst/officeforge/xlsx"
)

func handleXlsxSingle(args []string) {
	if len(args) < 8 {
		fmt.Println("Error: Missing required arguments")
		fmt.Println("\nUsage:")
		fmt.Println("  officeforge xlsx-single --input <template> --output <file> --key <keyword> --value <replacement>")
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

	keyword = internal.NormalizeKey(keyword)

	err := xlsx.ProcessXlsxSingle(inputPath, outputPath, keyword, replacement)
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("✓ Spreadsheet created: %s\n", outputPath)
}

func handleXlsxMulti(args []string) {
	if len(args) < 6 {
		fmt.Println("Error: Missing required arguments")
		fmt.Println("\nUsage:")
		fmt.Println("  officeforge xlsx-multi --input <template> --output <file> --data <json_file>")
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

	replacements = internal.NormalizeReplacements(replacements)

	err = xlsx.ProcessXlsxMulti(inputPath, outputPath, replacements)
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("✓ Spreadsheet created: %s\n", outputPath)
	fmt.Printf("  Replaced %d keywords\n", len(replacements))
}

func handleXlsxBatch(args []string) {
	if len(args) < 6 {
		fmt.Println("Error: Missing required arguments")
		fmt.Println("\nUsage:")
		fmt.Println("  officeforge xlsx-batch --input <template> --output <directory> --data <csv_or_json_file> [--pattern <pattern>]")
		fmt.Println("\nPattern examples:")
		fmt.Printf("  --pattern \"report_%%d.xlsx\"             Sequential: report_1.xlsx, report_2.xlsx\n")
		fmt.Println("  --pattern \"{NAME}_report.xlsx\"        From data: Alice_report.xlsx, Bob_report.xlsx")
		fmt.Println("  --pattern \"{ID}_{COMPANY}.xlsx\"       Multiple fields: 001_Acme.xlsx, 002_TechCorp.xlsx")
		fmt.Println("  --pattern \"{NAME}_{INDEX}.xlsx\"       Combine data and index: Alice_1.xlsx, Bob_2.xlsx")
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

	normalizedRecords := make([]map[string]string, len(records))
	for i, record := range records {
		normalizedRecords[i] = internal.NormalizeReplacements(record)
	}

	// Validate pattern if provided
	if pattern != "" {
		if err := internal.ValidatePattern(pattern, normalizedRecords[0]); err != nil {
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

	// Process spreadsheets using the pattern
	err = xlsx.ProcessXlsxMultipleRecords(inputPath, outputDir, normalizedRecords, pattern)
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("✓ Generated %d spreadsheets in: %s\n", len(records), outputDir)
}

func handleXlsxCheck(args []string) {
	if len(args) < 4 {
		fmt.Println("Error: Missing required arguments")
		fmt.Println("\nUsage:")
		fmt.Println("  officeforge xlsx-check --input <file> [--keys \"K1,K2\" | --data <json_file>] [--json]")
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

	// Logic: Error if input is missing OR (both keys AND data are missing)
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

		// Check for CSV support via file extension
		if strings.HasSuffix(strings.ToLower(dataPath), ".csv") {
			reader := csv.NewReader(bytes.NewReader(data))
			headers, err := reader.Read() // Reads the first row (the headers)
			if err != nil {
				fmt.Printf("Error parsing CSV headers: %v\n", err)
				os.Exit(1)
			}
			keywords = headers
		} else {
			// JSON Logic: Try parsing as an array of objects (starts with [)
			var dataList []map[string]any
			if err := json.Unmarshal(data, &dataList); err == nil {
				if len(dataList) > 0 {
					for k := range dataList[0] {
						keywords = append(keywords, k)
					}
				}
			} else {
				// Try parsing as a single object (starts with {)
				var keyMap map[string]any
				if err := json.Unmarshal(data, &keyMap); err == nil {
					for k := range keyMap {
						keywords = append(keywords, k)
					}
				} else {
					// Try parsing as a simple array of strings (starts with ["str"])
					var keyList []string
					if err := json.Unmarshal(data, &keyList); err == nil {
						keywords = keyList
					}
				}
			}
		}

		// Validation to ensure keywords were actually extracted
		if len(keywords) == 0 {
			fmt.Println("Error: No keywords found. Check if your CSV has headers or if your JSON is formatted correctly.")
			os.Exit(1)
		}
	} else {
		// Existing logic for comma-separated --keys flag
		keywords = strings.Split(keysString, ",")
		for i := range keywords {
			keywords[i] = strings.TrimSpace(keywords[i])
		}
	}

	for i := range keywords {
		keywords[i] = internal.NormalizeKey(keywords[i])
	}

	// Call the validation function
	results, err := internal.ValidateXlsxKeywords(inputPath, keywords)
	if err != nil {
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

	// Output results
	if outputJson {
		jsonBytes, _ := json.Marshal(results)
		fmt.Println(string(jsonBytes))
	} else {
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
