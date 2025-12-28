package internal

import (
	"archive/zip"
	"regexp"
	"strings"
)

// ProcessXlsxZipFile processes a single file from an XLSX zip archive.
func ProcessXlsxZipFile(file *zip.File, zipWriter *zip.Writer, replacements map[string]string) error {
	return ProcessZipFile(file, zipWriter, func(fileName string, content []byte) []byte {
		// Process the shared strings XML file where most text is stored
		if fileName == "xl/sharedStrings.xml" {
			processedContent := processSharedStringsXML(string(content), replacements)
			return []byte(processedContent)
		}
		// For other files, copy as-is
		return content
	})
}

func processSharedStringsXML(xmlContent string, replacements map[string]string) string {
	stringItems := splitIntoStringItems(xmlContent)

	for i, item := range stringItems {
		if strings.Contains(item, "<t>") || strings.Contains(item, "<t ") {
			plainText := extractTextFromStringItem(item)

			if ContainsAnyKeyword(plainText, replacements) {
				positionMap := buildStringItemPositionMap(item)
				stringItems[i] = ApplyReplacements(item, plainText, replacements, positionMap)
			}
		}
	}
	return strings.Join(stringItems, "")
}

func splitIntoStringItems(xmlContent string) []string {
	// Match <si> elements (string items) in the shared strings table
	re := regexp.MustCompile(`(<si\b[^>]*>(?:.*?)</si>)`)
	var result []string
	lastEnd := 0

	for _, item := range re.FindAllStringIndex(xmlContent, -1) {
		if item[0] > lastEnd {
			result = append(result, xmlContent[lastEnd:item[0]])
		}

		result = append(result, xmlContent[item[0]:item[1]])
		lastEnd = item[1]
	}

	if lastEnd < len(xmlContent) {
		result = append(result, xmlContent[lastEnd:])
	}

	return result
}

func extractTextFromStringItem(item string) string {
	// Extract text from <t> tags within string items
	// Note: XLSX can have <t> tags with attributes like xml:space="preserve"
	re := regexp.MustCompile(`<t(?:\s[^>]*)?>(.*?)</t>`)
	matches := re.FindAllStringSubmatch(item, -1)

	var text strings.Builder
	for _, match := range matches {
		if len(match) > 1 {
			text.WriteString(match[1])
		}
	}

	return text.String()
}

func buildStringItemPositionMap(item string) map[int]int {
	positionMap := make(map[int]int)

	// Build position map for <t> tags in string items
	re := regexp.MustCompile(`<t(?:\s[^>]*)?>(.*?)</t>`)
	matches := re.FindAllStringSubmatchIndex(item, -1)

	plainPos := 0
	for _, match := range matches {
		if len(match) >= 4 {
			contentStart := match[2] // Start of content group
			contentEnd := match[3]   // End of content group

			textContent := item[contentStart:contentEnd]
			for i := 0; i < len(textContent); i++ {
				positionMap[plainPos+i] = contentStart + i
			}
			plainPos += len(textContent)
		}
	}
	return positionMap
}
