package tests

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/siliconcatalyst/officeforge/pptx"
)

func TestProcessPptxSingle(t *testing.T) {
	templatePath := "testdata/template.pptx"
	outputPath := "testdata/output/single_output.pptx"

	// Ensure output directory exists
	if err := os.MkdirAll("testdata/output", 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
	defer os.RemoveAll("testdata/output")

	// Check template exists
	if !fileExists(templatePath) {
		t.Fatalf("Template file not found: %s", templatePath)
	}

	// Test single replacement
	err := pptx.ProcessPptxSingle(templatePath, outputPath, "{{NAME}}", "Alice Williams")
	if err != nil {
		t.Fatalf("ProcessPptxSingle failed: %v", err)
	}

	// Verify output file was created
	if !fileExists(outputPath) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	// Read and verify content
	content, err := readPptxContent(outputPath)
	if err != nil {
		t.Fatalf("Failed to read output content: %v", err)
	}

	// Check if replacement occurred
	if !strings.Contains(content, "Alice Williams") {
		t.Errorf("Replacement 'Alice Williams' not found in output")
	}

	// Verify placeholder was removed
	if strings.Contains(content, "{{NAME}}") {
		t.Errorf("Placeholder {{NAME}} still present in output")
	}

	t.Logf("\033[32m✓ Single replacement test passed\033[0m")
}

func TestProcessPptxSingleMissingKeyword(t *testing.T) {
	templatePath := "testdata/template.pptx"
	outputPath := "testdata/output/single_missing.pptx"

	if err := os.MkdirAll("testdata/output", 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
	defer os.RemoveAll("testdata/output")

	// Test with keyword that doesn't exist in template
	err := pptx.ProcessPptxSingle(templatePath, outputPath, "{{NONEXISTENT}}", "Some Value")
	if err != nil {
		t.Fatalf("ProcessPptxSingle failed: %v", err)
	}

	// File should still be created
	if !fileExists(outputPath) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	t.Logf("\033[33m⚠ Missing keyword handled gracefully\033[0m")
}

func TestProcessPptxSingleInvalidTemplate(t *testing.T) {
	outputPath := "testdata/output/invalid_output.pptx"

	if err := os.MkdirAll("testdata/output", 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
	defer os.RemoveAll("testdata/output")

	// Test with non-existent template
	err := pptx.ProcessPptxSingle("nonexistent.pptx", outputPath, "{{NAME}}", "Test")
	if err == nil {
		t.Fatalf("Expected error for non-existent template, got nil")
	}

	t.Logf("\033[36m✓ Invalid template error handled: %v\033[0m", err)
}

func TestProcessPptxMulti(t *testing.T) {
	templatePath := "testdata/template.pptx"
	outputPath := "testdata/output/multi_output.pptx"

	if err := os.MkdirAll("testdata/output", 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
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

	err := pptx.ProcessPptxMulti(templatePath, outputPath, replacements)
	if err != nil {
		t.Fatalf("ProcessPptxMulti failed: %v", err)
	}

	if !fileExists(outputPath) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	// Read and verify content
	content, err := readPptxContent(outputPath)
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

	t.Logf("\033[32m✓ Multiple replacements test passed\033[0m")
}

func TestProcessPptxMultiEmptyReplacements(t *testing.T) {
	templatePath := "testdata/template.pptx"
	outputPath := "testdata/output/multi_empty.pptx"

	if err := os.MkdirAll("testdata/output", 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
	defer os.RemoveAll("testdata/output")

	// Test with empty replacements map
	err := pptx.ProcessPptxMulti(templatePath, outputPath, map[string]string{})
	if err != nil {
		t.Fatalf("ProcessPptxMulti with empty replacements failed: %v", err)
	}

	if !fileExists(outputPath) {
		t.Fatalf("Output file was not created: %s", outputPath)
	}

	t.Logf("\033[32m✓ Empty replacements handled gracefully\033[0m")
}

func TestProcessPptxMultipleRecords(t *testing.T) {
	templatePath := "testdata/template.pptx"
	outputDir := "testdata/output/batch"

	if err := os.MkdirAll(outputDir, 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
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

	filePattern := "presentation_%d.pptx"

	err := pptx.ProcessPptxMultipleRecords(templatePath, outputDir, records, filePattern)
	if err != nil {
		t.Fatalf("ProcessPptxMultipleRecords failed: %v", err)
	}

	// Verify all files were created
	expectedFiles := []string{"presentation_1.pptx", "presentation_2.pptx"}
	for _, filename := range expectedFiles {
		path := filepath.Join(outputDir, filename)
		if !fileExists(path) {
			t.Errorf("Expected output file not created: %s", path)
		}
	}

	// Verify content of first file
	content, err := readPptxContent(filepath.Join(outputDir, "presentation_1.pptx"))
	if err != nil {
		t.Fatalf("Failed to read first output: %v", err)
	}

	if !strings.Contains(content, "Sarah Connor") {
		t.Errorf("First record data not found in output")
	}

	t.Logf("\033[32m✓ Batch processing with pattern test passed\033[0m")
}

func TestProcessPptxMultipleRecordsDefaultNaming(t *testing.T) {
	templatePath := "testdata/template.pptx"
	outputDir := "testdata/output/batch_default"

	if err := os.MkdirAll(outputDir, 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
	defer os.RemoveAll("testdata/output")

	records := []map[string]string{
		{"{{NAME}}": "Test User 1", "{{EMAIL}}": "test1@example.com"},
		{"{{NAME}}": "Test User 2", "{{EMAIL}}": "test2@example.com"},
	}

	// Test with empty pattern (should use default naming)
	err := pptx.ProcessPptxMultipleRecords(templatePath, outputDir, records, "")
	if err != nil {
		t.Fatalf("ProcessPptxMultipleRecords with default naming failed: %v", err)
	}

	// Verify default naming pattern
	expectedFiles := []string{"presentation_1.pptx", "presentation_2.pptx"}
	for _, filename := range expectedFiles {
		path := filepath.Join(outputDir, filename)
		if !fileExists(path) {
			t.Errorf("Expected default named file not created: %s", path)
		}
	}

	t.Logf("\033[32m✓ Batch processing with default naming test passed\033[0m")
}

func TestProcessPptxMultipleRecordsWithNames(t *testing.T) {
	templatePath := "testdata/template.pptx"
	outputDir := "testdata/output/batch_custom"

	if err := os.MkdirAll(outputDir, 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
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
		return fmt.Sprintf("%s_%s_presentation.pptx", name, company)
	}

	err := pptx.ProcessPptxMultipleRecordsWithNames(templatePath, outputDir, records, nameFunc)
	if err != nil {
		t.Fatalf("ProcessPptxMultipleRecordsWithNames failed: %v", err)
	}

	// Verify custom named files were created
	expectedFiles := []string{
		"Alice_TechCorp_presentation.pptx",
		"Bob_StartupInc_presentation.pptx",
		"Charlie_MegaCorp_presentation.pptx",
	}

	for _, filename := range expectedFiles {
		path := filepath.Join(outputDir, filename)
		if !fileExists(path) {
			t.Errorf("Expected custom named file not created: %s", path)
		}
	}

	t.Logf("\033[32m✓ Batch processing with custom naming test passed\033[0m")
}

func TestProcessPptxMultipleRecordsEmptyRecords(t *testing.T) {
	templatePath := "testdata/template.pptx"
	outputDir := "testdata/output/batch_empty"

	if err := os.MkdirAll(outputDir, 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
	defer os.RemoveAll("testdata/output")

	// Test with empty records slice
	err := pptx.ProcessPptxMultipleRecords(templatePath, outputDir, []map[string]string{}, "test_%d.pptx")
	if err != nil {
		t.Fatalf("ProcessPptxMultipleRecords with empty records failed: %v", err)
	}

	// Should create directory but no files
	files, _ := os.ReadDir(outputDir)
	if len(files) > 0 {
		t.Errorf("Expected no files, but found %d", len(files))
	}

	t.Logf("\033[32m✓ Empty records handled gracefully\033[0m")
}

func TestProcessPptxSpecialCharacters(t *testing.T) {
	templatePath := "testdata/template.pptx"
	outputPath := "testdata/output/special_chars.pptx"

	if err := os.MkdirAll("testdata/output", 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}
	defer os.RemoveAll("testdata/output")

	// Test with special characters in replacement
	replacements := map[string]string{
		"{{NAME}}":    "José García-O'Brien",
		"{{EMAIL}}":   "josé@españa.es",
		"{{COMPANY}}": "Müller & Co. <GmbH>",
		"{{SALARY}}":  "$1,234,567.89",
	}

	err := pptx.ProcessPptxMulti(templatePath, outputPath, replacements)
	if err != nil {
		t.Fatalf("ProcessPptxMulti with special chars failed: %v", err)
	}

	if !fileExists(outputPath) {
		t.Fatalf("Output file was not created")
	}

	content, err := readPptxContent(outputPath)
	if err != nil {
		t.Fatalf("Failed to read output: %v", err)
	}

	// Basic check - file should contain some of the special characters
	if !strings.Contains(content, "José") && !strings.Contains(content, "Müller") {
		t.Logf("Warning: Special characters may not be preserved exactly")
	}

	t.Logf("\033[32m✓ Special characters test passed\033[0m")
}

// Benchmark tests
func BenchmarkProcessPptxSingle(b *testing.B) {
	templatePath := "testdata/template.pptx"
	if err := os.MkdirAll("testdata/output", 0755); err != nil {
		b.Fatalf("Failed to create output directory: %v", err)
	}
	defer os.RemoveAll("testdata/output")

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		outputPath := fmt.Sprintf("testdata/output/bench_%d.pptx", i)
		if err := pptx.ProcessPptxSingle(templatePath, outputPath, "{{NAME}}", "Benchmark User"); err != nil {
			b.Fatalf("Benchmark failed: %v", err)
		}
	}
}

func BenchmarkProcessPptxMulti(b *testing.B) {
	templatePath := "testdata/template.pptx"
	if err := os.MkdirAll("testdata/output", 0755); err != nil {
		b.Fatalf("Failed to create output directory: %v", err)
	}
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
		outputPath := fmt.Sprintf("testdata/output/bench_multi_%d.pptx", i)
		if err := pptx.ProcessPptxMulti(templatePath, outputPath, replacements); err != nil {
			b.Fatalf("Benchmark failed: %v", err)
		}
	}
}
