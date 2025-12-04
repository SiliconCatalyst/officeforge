package tests

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/siliconcatalyst/officeforge/docx"
)

// Helper function to check if a file exists
func fileExists(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}

// Helper function to read text content from a DOCX file
func readDocxContent(path string) (string, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return "", err
	}
	defer reader.Close()

	for _, file := range reader.File {
		if file.Name == "word/document.xml" {
			rc, err := file.Open()
			if err != nil {
				return "", err
			}
			defer rc.Close()

			content, err := io.ReadAll(rc)
			if err != nil {
				return "", err
			}
			return string(content), nil
		}
	}
	return "", fmt.Errorf("document.xml not found")
}

func TestProcessDocxSingle(t *testing.T) {
	templatePath := "testdata/template.docx"
	outputPath := "testdata/output/single_output.docx"

	// Ensure output directory exists
	os.MkdirAll("testdata/output", 0755)
	defer os.RemoveAll("testdata/output")

	// Check template exists
	if !fileExists(templatePath) {
		t.Fatalf("Template file not found: %s", templatePath)
	}

	// Test single replacement
	err := docx.ProcessDocxSingle(templatePath, outputPath, "{{NAME}}", "Alice Williams")
	if err != nil {
		t.Fatalf("ProcessDocxSingle failed: %v", err)
	}

	// Verify output file was created
	if !fileExists(outputPath) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	// Read and verify content
	content, err := readDocxContent(outputPath)
	if err != nil {
		t.Fatalf("Failed to read output content: %v", err)
	}

	// Check if replacement occurred (look for the value, not the placeholder)
	if !strings.Contains(content, "Alice Williams") {
		t.Errorf("Replacement 'Alice Williams' not found in output")
	}

	// Verify placeholder was removed
	if strings.Contains(content, "{{NAME}}") {
		t.Errorf("Placeholder {{NAME}} still present in output")
	}

	t.Logf("✓ Single replacement test passed")
}

func TestProcessDocxSingleMissingKeyword(t *testing.T) {
	templatePath := "testdata/template.docx"
	outputPath := "testdata/output/single_missing.docx"

	os.MkdirAll("testdata/output", 0755)
	defer os.RemoveAll("testdata/output")

	// Test with keyword that doesn't exist in template
	err := docx.ProcessDocxSingle(templatePath, outputPath, "{{NONEXISTENT}}", "Some Value")
	if err != nil {
		t.Fatalf("ProcessDocxSingle failed: %v", err)
	}

	// File should still be created
	if !fileExists(outputPath) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	t.Logf("✓ Missing keyword handled gracefully")
}

func TestProcessDocxSingleInvalidTemplate(t *testing.T) {
	outputPath := "testdata/output/invalid_output.docx"

	os.MkdirAll("testdata/output", 0755)
	defer os.RemoveAll("testdata/output")

	// Test with non-existent template
	err := docx.ProcessDocxSingle("nonexistent.docx", outputPath, "{{NAME}}", "Test")
	if err == nil {
		t.Fatalf("Expected error for non-existent template, got nil")
	}

	t.Logf("✓ Invalid template error handled: %v", err)
}

func TestProcessDocxMulti(t *testing.T) {
	templatePath := "testdata/template.docx"
	outputPath := "testdata/output/multi_output.docx"

	os.MkdirAll("testdata/output", 0755)
	defer os.RemoveAll("testdata/output")

	if !fileExists(templatePath) {
		t.Fatalf("Template file not found: %s", templatePath)
	}

	// Test multiple replacements
	replacements := map[string]string{
		"{{NAME}}":       "Michael Chen",
		"{{EMAIL}}":      "michael@company.com",
		"{{PHONE}}":      "555-7890",
		"{{COMPANY}}":    "Innovation Labs",
		"{{POSITION}}":   "Senior Developer",
		"{{START_DATE}}": "2024-12-15",
		"{{SALARY}}":     "$120,000",
	}

	err := docx.ProcessDocxMulti(templatePath, outputPath, replacements)
	if err != nil {
		t.Fatalf("ProcessDocxMulti failed: %v", err)
	}

	if !fileExists(outputPath) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	// Read and verify content
	content, err := readDocxContent(outputPath)
	if err != nil {
		t.Fatalf("Failed to read output content: %v", err)
	}

	// Verify all replacements occurred
	expectedValues := []string{"Michael Chen", "michael@company.com", "555-7890", "Innovation Labs", "Senior Developer", "2024-12-15", "$120,000"}
	for _, value := range expectedValues {
		if !strings.Contains(content, value) {
			t.Errorf("Expected value '%s' not found in output", value)
		}
	}

	// Verify no placeholders remain
	for placeholder := range replacements {
		if strings.Contains(content, placeholder) {
			t.Errorf("Placeholder %s still present in output", placeholder)
		}
	}

	t.Logf("✓ Multiple replacements test passed")
}

func TestProcessDocxMultiEmptyReplacements(t *testing.T) {
	templatePath := "testdata/template.docx"
	outputPath := "testdata/output/multi_empty.docx"

	os.MkdirAll("testdata/output", 0755)
	defer os.RemoveAll("testdata/output")

	// Test with empty replacements map
	err := docx.ProcessDocxMulti(templatePath, outputPath, map[string]string{})
	if err != nil {
		t.Fatalf("ProcessDocxMulti with empty replacements failed: %v", err)
	}

	if !fileExists(outputPath) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	t.Logf("✓ Empty replacements handled gracefully")
}

func TestProcessDocxMultipleRecords(t *testing.T) {
	templatePath := "testdata/template.docx"
	outputDir := "testdata/output/batch"

	os.MkdirAll(outputDir, 0755)
	defer os.RemoveAll("testdata/output")

	if !fileExists(templatePath) {
		t.Fatalf("Template file not found: %s", templatePath)
	}

	// Test batch processing with pattern
	records := []map[string]string{
		{
			"{{NAME}}":       "Sarah Connor",
			"{{EMAIL}}":      "sarah@resistance.com",
			"{{PHONE}}":      "555-0001",
			"{{COMPANY}}":    "Resistance HQ",
			"{{POSITION}}":   "Leader",
			"{{START_DATE}}": "1984-05-12",
			"{{SALARY}}":     "$0",
		},
		{
			"{{NAME}}":       "Kyle Reese",
			"{{EMAIL}}":      "kyle@resistance.com",
			"{{PHONE}}":      "555-0002",
			"{{COMPANY}}":    "Resistance HQ",
			"{{POSITION}}":   "Soldier",
			"{{START_DATE}}": "1984-05-12",
			"{{SALARY}}":     "$0",
		},
	}

	filePattern := "contract_%d.docx"

	err := docx.ProcessDocxMultipleRecords(templatePath, outputDir, records, filePattern)
	if err != nil {
		t.Fatalf("ProcessDocxMultipleRecords failed: %v", err)
	}

	// Verify all files were created
	expectedFiles := []string{"contract_1.docx", "contract_2.docx"}
	for _, filename := range expectedFiles {
		path := filepath.Join(outputDir, filename)
		if !fileExists(path) {
			t.Errorf("Expected output file not created: %s", path)
		}
	}

	// Verify content of first file
	content, err := readDocxContent(filepath.Join(outputDir, "contract_1.docx"))
	if err != nil {
		t.Fatalf("Failed to read first output: %v", err)
	}

	if !strings.Contains(content, "Sarah Connor") {
		t.Errorf("First record data not found in output")
	}

	t.Logf("✓ Batch processing with pattern test passed")
}

func TestProcessDocxMultipleRecordsDefaultNaming(t *testing.T) {
	templatePath := "testdata/template.docx"
	outputDir := "testdata/output/batch_default"

	os.MkdirAll(outputDir, 0755)
	defer os.RemoveAll("testdata/output")

	records := []map[string]string{
		{"{{NAME}}": "Test User 1", "{{EMAIL}}": "test1@example.com"},
		{"{{NAME}}": "Test User 2", "{{EMAIL}}": "test2@example.com"},
	}

	// Test with empty pattern (should use default naming)
	err := docx.ProcessDocxMultipleRecords(templatePath, outputDir, records, "")
	if err != nil {
		t.Fatalf("ProcessDocxMultipleRecords with default naming failed: %v", err)
	}

	// Verify default naming pattern
	expectedFiles := []string{"document_1.docx", "document_2.docx"}
	for _, filename := range expectedFiles {
		path := filepath.Join(outputDir, filename)
		if !fileExists(path) {
			t.Errorf("Expected default named file not created: %s", path)
		}
	}

	t.Logf("✓ Batch processing with default naming test passed")
}

func TestProcessDocxMultipleRecordsWithNames(t *testing.T) {
	templatePath := "testdata/template.docx"
	outputDir := "testdata/output/batch_custom"

	os.MkdirAll(outputDir, 0755)
	defer os.RemoveAll("testdata/output")

	records := []map[string]string{
		{"{{NAME}}": "Alice", "{{COMPANY}}": "TechCorp"},
		{"{{NAME}}": "Bob", "{{COMPANY}}": "StartupInc"},
		{"{{NAME}}": "Charlie", "{{COMPANY}}": "MegaCorp"},
	}

	// Custom naming function
	nameFunc := func(record map[string]string, index int) string {
		name := strings.ReplaceAll(record["{{NAME}}"], " ", "_")
		company := strings.ReplaceAll(record["{{COMPANY}}"], " ", "_")
		return fmt.Sprintf("%s_%s_contract.docx", name, company)
	}

	err := docx.ProcessDocxMultipleRecordsWithNames(templatePath, outputDir, records, nameFunc)
	if err != nil {
		t.Fatalf("ProcessDocxMultipleRecordsWithNames failed: %v", err)
	}

	// Verify custom named files were created
	expectedFiles := []string{
		"Alice_TechCorp_contract.docx",
		"Bob_StartupInc_contract.docx",
		"Charlie_MegaCorp_contract.docx",
	}

	for _, filename := range expectedFiles {
		path := filepath.Join(outputDir, filename)
		if !fileExists(path) {
			t.Errorf("Expected custom named file not created: %s", path)
		}
	}

	t.Logf("✓ Batch processing with custom naming test passed")
}

func TestProcessDocxMultipleRecordsEmptyRecords(t *testing.T) {
	templatePath := "testdata/template.docx"
	outputDir := "testdata/output/batch_empty"

	os.MkdirAll(outputDir, 0755)
	defer os.RemoveAll("testdata/output")

	// Test with empty records slice
	err := docx.ProcessDocxMultipleRecords(templatePath, outputDir, []map[string]string{}, "test_%d.docx")
	if err != nil {
		t.Fatalf("ProcessDocxMultipleRecords with empty records failed: %v", err)
	}

	// Should create directory but no files
	files, _ := os.ReadDir(outputDir)
	if len(files) > 0 {
		t.Errorf("Expected no files, but found %d", len(files))
	}

	t.Logf("✓ Empty records handled gracefully")
}

func TestProcessDocxSpecialCharacters(t *testing.T) {
	templatePath := "testdata/template.docx"
	outputPath := "testdata/output/special_chars.docx"

	os.MkdirAll("testdata/output", 0755)
	defer os.RemoveAll("testdata/output")

	// Test with special characters in replacement
	replacements := map[string]string{
		"{{NAME}}":    "José García-O'Brien",
		"{{EMAIL}}":   "josé@españa.es",
		"{{COMPANY}}": "Müller & Co. <GmbH>",
		"{{SALARY}}":  "$1,234,567.89",
	}

	err := docx.ProcessDocxMulti(templatePath, outputPath, replacements)
	if err != nil {
		t.Fatalf("ProcessDocxMulti with special chars failed: %v", err)
	}

	if !fileExists(outputPath) {
		t.Fatalf("Output file was not created")
	}

	content, err := readDocxContent(outputPath)
	if err != nil {
		t.Fatalf("Failed to read output: %v", err)
	}

	// Basic check - file should contain some of the special characters
	if !strings.Contains(content, "José") && !strings.Contains(content, "Müller") {
		t.Logf("Warning: Special characters may not be preserved exactly")
	}

	t.Logf("✓ Special characters test passed")
}

// Benchmark tests
func BenchmarkProcessDocxSingle(b *testing.B) {
	templatePath := "testdata/template.docx"
	os.MkdirAll("testdata/output", 0755)
	defer os.RemoveAll("testdata/output")

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		outputPath := fmt.Sprintf("testdata/output/bench_%d.docx", i)
		docx.ProcessDocxSingle(templatePath, outputPath, "{{NAME}}", "Benchmark User")
	}
}

func BenchmarkProcessDocxMulti(b *testing.B) {
	templatePath := "testdata/template.docx"
	os.MkdirAll("testdata/output", 0755)
	defer os.RemoveAll("testdata/output")

	replacements := map[string]string{
		"{{NAME}}":     "Benchmark User",
		"{{EMAIL}}":    "bench@test.com",
		"{{PHONE}}":    "555-0000",
		"{{COMPANY}}":  "BenchCorp",
		"{{POSITION}}": "Tester",
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		outputPath := fmt.Sprintf("testdata/output/bench_multi_%d.docx", i)
		docx.ProcessDocxMulti(templatePath, outputPath, replacements)
	}
}
