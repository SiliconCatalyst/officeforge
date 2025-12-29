package internal

import (
	"archive/zip"
	"fmt"
	"io"
	"regexp"
	"strings"
)

// PatternType represents the type of naming pattern
type PatternType int

const (
	PatternTypeUnknown    PatternType = iota
	PatternTypeSequential             // Uses %d format (e.g., "file_%d.docx")
	PatternTypeData                   // Uses {{FIELD}} placeholders (e.g., "{{NAME}}_contract.docx")
)

// DetectPatternType determines what type of pattern string is being used
func DetectPatternType(pattern string) PatternType {
	if pattern == "" {
		return PatternTypeUnknown
	}

	hasPlaceholders := strings.Contains(pattern, "{") && strings.Contains(pattern, "}")
	hasFormatVerb := strings.Contains(pattern, "%d")

	if hasPlaceholders {
		return PatternTypeData
	}
	if hasFormatVerb {
		return PatternTypeSequential
	}

	return PatternTypeUnknown
}

// ReplacePlaceholders replaces {{FIELD}} placeholders with values from the record
// Supports {{FIELD}} for any record key and {{INDEX}} for the record number
func ReplacePlaceholders(pattern string, record map[string]string, index int) string {
	result := pattern

	// Replace all {KEY} placeholders with corresponding record values
	for key, value := range record {
		placeholder := fmt.Sprintf("%s", key)
		result = strings.ReplaceAll(result, placeholder, value)
	}

	// Support {INDEX} for the record number
	result = strings.ReplaceAll(result, "{{INDEX}}", fmt.Sprintf("%d", index))

	// Sanitize the filename
	result = SanitizeFilename(result)

	return result
}

// SanitizeFilename removes or replaces invalid filename characters
func SanitizeFilename(filename string) string {
	// Replace invalid characters with underscores
	invalid := []string{"/", "\\", ":", "*", "?", "\"", "<", ">", "|"}
	result := filename
	for _, char := range invalid {
		result = strings.ReplaceAll(result, char, "_")
	}

	// Remove any control characters
	result = strings.Map(func(r rune) rune {
		if r < 32 {
			return -1
		}
		return r
	}, result)

	// Trim whitespace
	result = strings.TrimSpace(result)

	return result
}

// ExtractPlaceholders returns all placeholder names found in a pattern
// Example: "{{NAME}}_{{ID}}.docx" returns ["NAME", "ID"]
func ExtractPlaceholders(pattern string) []string {
	re := regexp.MustCompile(`\{\{([^}]+)\}\}`)
	matches := re.FindAllStringSubmatch(pattern, -1)

	placeholders := make([]string, 0, len(matches))
	for _, match := range matches {
		if len(match) > 1 {
			// Return just the field name without braces
			placeholders = append(placeholders, "{{"+match[1]+"}}")
		}
	}
	return placeholders
}

// ValidatePattern checks if a pattern is valid and returns an error if not
func ValidatePattern(pattern string, sampleRecord map[string]string) error {
	if pattern == "" {
		return nil // Empty pattern is valid (will use default)
	}

	patternType := DetectPatternType(pattern)

	switch patternType {
	case PatternTypeSequential:
		// Just check that %d exists
		if !strings.Contains(pattern, "%d") {
			return fmt.Errorf("sequential pattern must contain %%d")
		}
		return nil

	case PatternTypeData:
		// Check that all placeholders exist in the sample record
		placeholders := ExtractPlaceholders(pattern)
		if len(placeholders) == 0 {
			return fmt.Errorf("data pattern must contain at least one {{FIELD}} placeholder")
		}

		// Validate placeholders against sample record (skip {{INDEX}} as it's built-in)
		for _, placeholder := range placeholders {
			if placeholder == "INDEX" {
				continue
			}
			if _, exists := sampleRecord[placeholder]; !exists {
				return fmt.Errorf("placeholder %s not found in data fields. Available fields: %v",
					placeholder, getMapKeys(sampleRecord))
			}
		}
		return nil

	case PatternTypeUnknown:
		return fmt.Errorf("pattern must contain: {{FIELD}}, {{INDEX}} placeholders or %%d for sequential numbering")
	}

	return nil
}

// CreateNamingFunction creates the appropriate naming function based on the pattern
func createNamingFunction(pattern, baseName, extension string) func(map[string]string, int) string {
	if pattern == "" {
		// Default naming using the provided base and extension
		return func(record map[string]string, index int) string {
			return fmt.Sprintf("%s_%d%s", baseName, index, extension)
		}
	}

	patternType := DetectPatternType(pattern)

	switch patternType {
	case PatternTypeSequential:
		return func(record map[string]string, index int) string {
			return fmt.Sprintf(pattern, index)
		}

	case PatternTypeData:
		return func(record map[string]string, index int) string {
			return ReplacePlaceholders(pattern, record, index)
		}

	default:
		// Fallback
		return func(record map[string]string, index int) string {
			return fmt.Sprintf("%s_%d%s", baseName, index, extension)
		}
	}
}

func CreateDocxNamingFunction(pattern string) func(map[string]string, int) string {
	return createNamingFunction(pattern, "document", ".docx")
}

func CreatePptxNamingFunction(pattern string) func(map[string]string, int) string {
	return createNamingFunction(pattern, "presentation", ".pptx")
}

func CreateXlsxNamingFunction(pattern string) func(map[string]string, int) string {
	return createNamingFunction(pattern, "spreadsheet", ".xlsx")
}

// getMapKeys returns all keys from a map as a slice (helper for error messages)
func getMapKeys(m map[string]string) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	return keys
}

func ValidateDocxKeywords(inputPath string, keywords []string) (map[string]bool, error) {
	reader, err := zip.OpenReader(inputPath)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	results := make(map[string]bool)
	for _, k := range keywords {
		results[k] = false
	}

	for _, file := range reader.File {
		// Optimization: only read text-heavy XML files
		if !strings.HasPrefix(file.Name, "word/") || !strings.HasSuffix(file.Name, ".xml") {
			continue
		}

		rc, err := file.Open()
		if err != nil {
			return nil, err
		}
		content, _ := io.ReadAll(rc)
		rc.Close()

		contentStr := string(content)
		for _, k := range keywords {
			if !results[k] && strings.Contains(contentStr, k) {
				results[k] = true
			}
		}
	}
	return results, nil
}

func ValidatePptxKeywords(inputPath string, keywords []string) (map[string]bool, error) {
	reader, err := zip.OpenReader(inputPath)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	results := make(map[string]bool)
	for _, k := range keywords {
		results[k] = false
	}

	for _, file := range reader.File {
		// PPTX slide content is inside the "ppt/slides/" folder
		if !strings.HasPrefix(file.Name, "ppt/slides/") || !strings.HasSuffix(file.Name, ".xml") {
			continue
		}

		rc, err := file.Open()
		if err != nil {
			return nil, err
		}
		content, _ := io.ReadAll(rc)
		rc.Close()

		contentStr := string(content)
		for _, k := range keywords {
			if !results[k] && strings.Contains(contentStr, k) {
				results[k] = true
			}
		}
	}
	return results, nil
}

func ValidateXlsxKeywords(inputPath string, keywords []string) (map[string]bool, error) {
	reader, err := zip.OpenReader(inputPath)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	results := make(map[string]bool)
	for _, k := range keywords {
		results[k] = false
	}

	for _, file := range reader.File {
		// Check both individual worksheets and the shared strings table
		isWorksheet := strings.HasPrefix(file.Name, "xl/worksheets/") && strings.HasSuffix(file.Name, ".xml")
		isSharedStrings := file.Name == "xl/sharedStrings.xml"

		if !isWorksheet && !isSharedStrings {
			continue
		}

		rc, err := file.Open()
		if err != nil {
			return nil, err
		}
		content, _ := io.ReadAll(rc)
		rc.Close()

		contentStr := string(content)
		for _, k := range keywords {
			if !results[k] && strings.Contains(contentStr, k) {
				results[k] = true
			}
		}
	}
	return results, nil
}

// NormalizeKey ensures a key has double braces
// "CLIENT_NAME" -> "{{CLIENT_NAME}}"
// "{{CLIENT_NAME}}" -> "{{CLIENT_NAME}}"
func NormalizeKey(key string) string {
	// Already has braces
	if strings.HasPrefix(key, "{{") && strings.HasSuffix(key, "}}") {
		return key
	}
	// Add braces
	return fmt.Sprintf("{{%s}}", key)
}

// NormalizeReplacements converts all keys in a map to have double braces
// This allows users to provide data with or without braces
func NormalizeReplacements(replacements map[string]string) map[string]string {
	normalized := make(map[string]string, len(replacements))
	for key, value := range replacements {
		normalized[NormalizeKey(key)] = value
	}
	return normalized
}
