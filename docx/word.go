package docx

import (
	"archive/zip"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/siliconcatalyst/officeforge/internal"
)

func ProcessDocxSingle(inputPath, outputPath, keyword, replacement string) error {
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

	// Process each file in the DOCX
	for _, file := range reader.File {
		err := internal.ProcessZipFile(file, zipWriter, replacements)
		if err != nil {
			return fmt.Errorf("failed to process file %s: %v", file.Name, err)
		}
	}

	return nil
}

func ProcessDocxMulti(inputPath, outputPath string, replacements map[string]string) error {
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

	// Process each file in the DOCX
	for _, file := range reader.File {
		err := internal.ProcessZipFile(file, zipWriter, replacements)
		if err != nil {
			return fmt.Errorf("failed to process file %s: %v", file.Name, err)
		}
	}

	log.Printf("Successfully processed %d replacements in %s", len(replacements), outputPath)
	return nil
}

func ProcessDocxMultipleRecords(inputPath, outputDir string, records []map[string]string, fileNamePattern string) error {
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("failed to create output directory: %v", err)
	}

	for i, record := range records {
		// Generate output filename
		var outputPath string
		if fileNamePattern != "" {
			// Use pattern with record index
			outputPath = filepath.Join(outputDir, fmt.Sprintf(fileNamePattern, i+1))
		} else {
			// Default naming pattern
			outputPath = filepath.Join(outputDir, fmt.Sprintf("document_%d.docx", i+1))
		}

		// Process the document with this record's replacements
		err := ProcessDocxMulti(inputPath, outputPath, record)
		if err != nil {
			log.Printf("Failed to process record %d: %v", i+1, err)
			continue
		}

		log.Printf("Successfully created: %s", outputPath)
	}

	log.Printf("Successfully processed %d records", len(records))
	return nil
}

func ProcessDocxMultipleRecordsWithNames(inputPath, outputDir string, records []map[string]string, nameFunc func(map[string]string, int) string) error {
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("failed to create output directory: %v", err)
	}

	for i, record := range records {
		fileName := nameFunc(record, i+1)
		outputPath := filepath.Join(outputDir, fileName)

		err := ProcessDocxMulti(inputPath, outputPath, record)
		if err != nil {
			log.Printf("Failed to process record %d: %v", i+1, err)
			continue
		}

		log.Printf("Successfully created: %s", outputPath)
	}

	log.Printf("Successfully processed %d records", len(records))
	return nil
}
