package docx

import (
	"archive/zip"
	"fmt"
	"log"
	"os"

	"github.com/siliconcatalyst/officeforge/internal"
)

func ProcessDocxSingle(inputPath, outputPath, keyword, replacement string) error {
	// Read the input DOCX file
	reader, err := zip.OpenReader(inputPath)
	if err != nil {
		return fmt.Errorf("failed to open input file: %v", err)
	}
	defer reader.Close()

	// Create output DOCX file
	outputFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("failed to create output file: %v", err)
	}
	defer outputFile.Close()

	// Create zip writer for output
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

func ProcessDocxBatch(inputPath, outputPath string, replacements map[string]string) error {
	// Read the input DOCX file
	reader, err := zip.OpenReader(inputPath)
	if err != nil {
		return fmt.Errorf("failed to open input file: %v", err)
	}
	defer reader.Close()

	// Create output DOCX file
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
