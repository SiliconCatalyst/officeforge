package xlsx

import (
	"archive/zip"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/siliconcatalyst/officeforge/internal"
)

// ProcessXlsxSingle performs a single keyword replacement in an XLSX file
func ProcessXlsxSingle(inputPath, outputPath, keyword, replacement string) error {
	reader, err := zip.OpenReader(inputPath)
	if err != nil {
		return fmt.Errorf("failed to open input file: %v", err)
	}
	defer reader.Close()

	outputFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("failed to create output file: %v", err)
	}
	defer outputFile.Close()

	zipWriter := zip.NewWriter(outputFile)
	defer zipWriter.Close()

	// Create replacements map
	replacements := map[string]string{keyword: replacement}

	// Process each file in the XLSX
	for _, file := range reader.File {
		err := internal.ProcessXlsxZipFile(file, zipWriter, replacements)
		if err != nil {
			return fmt.Errorf("failed to process file %s: %v", file.Name, err)
		}
	}

	return nil
}

// ProcessXlsxMulti performs multiple keyword replacements in an XLSX file
func ProcessXlsxMulti(inputPath, outputPath string, replacements map[string]string) error {
	reader, err := zip.OpenReader(inputPath)
	if err != nil {
		return fmt.Errorf("failed to open input file: %v", err)
	}
	defer reader.Close()

	outputFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("failed to create output file: %v", err)
	}
	defer outputFile.Close()

	// Create zip writer for output
	zipWriter := zip.NewWriter(outputFile)
	defer zipWriter.Close()

	// Process each file in the XLSX
	for _, file := range reader.File {
		err := internal.ProcessXlsxZipFile(file, zipWriter, replacements)
		if err != nil {
			return fmt.Errorf("failed to process file %s: %v", file.Name, err)
		}
	}

	log.Printf("Successfully processed %d replacements in %s", len(replacements), outputPath)
	return nil
}

// ProcessXlsxMultipleRecords generates multiple XLSX files using a naming pattern
// Pattern can be:
//   - Sequential: "report_%d.xlsx" (uses index)
//   - Data-based: "{EMPLOYEE}_report.xlsx" (uses record fields)
//   - Empty: defaults to "spreadsheet_%d.xlsx"
func ProcessXlsxMultipleRecords(inputPath, outputDir string, records []map[string]string, fileNamePattern string) error {
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("failed to create output directory: %v", err)
	}

	// Validate pattern with first record
	if len(records) > 0 && fileNamePattern != "" {
		if err := internal.ValidatePattern(fileNamePattern, records[0]); err != nil {
			return fmt.Errorf("invalid pattern: %v", err)
		}
	}

	// Create naming function based on pattern
	nameFunc := internal.CreateNamingFunction(fileNamePattern)

	for i, record := range records {
		// Generate filename using the naming function
		fileName := nameFunc(record, i+1)
		outputPath := filepath.Join(outputDir, fileName)

		// Process the spreadsheet with this record's replacements
		err := ProcessXlsxMulti(inputPath, outputPath, record)
		if err != nil {
			log.Printf("Failed to process record %d: %v", i+1, err)
			continue
		}

		log.Printf("Successfully created: %s", outputPath)
	}

	log.Printf("Successfully processed %d records", len(records))
	return nil
}

// ProcessXlsxMultipleRecordsWithNames generates multiple XLSX files using a custom naming function
// This provides maximum flexibility for complex naming logic
func ProcessXlsxMultipleRecordsWithNames(inputPath, outputDir string, records []map[string]string, nameFunc func(map[string]string, int) string) error {
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("failed to create output directory: %v", err)
	}

	for i, record := range records {
		fileName := nameFunc(record, i+1)
		outputPath := filepath.Join(outputDir, fileName)

		err := ProcessXlsxMulti(inputPath, outputPath, record)
		if err != nil {
			log.Printf("Failed to process record %d: %v", i+1, err)
			continue
		}

		log.Printf("Successfully created: %s", outputPath)
	}

	log.Printf("Successfully processed %d records", len(records))
	return nil
}
