package internal

import (
	"archive/zip"
	"regexp"
	"strings"
)

// ProcessDocxZipFile processes a single file from a DOCX zip archive.
func ProcessDocxZipFile(file *zip.File, zipWriter *zip.Writer, replacements map[string]string) error {
	return ProcessZipFile(file, zipWriter, func(fileName string, content []byte) []byte {
		// If this is the main document XML file, process it for replacements
		if fileName == "word/document.xml" {
			processedContent := processDocumentXML(string(content), replacements)
			return []byte(processedContent)
		}
		// For other files, copy as-is
		return content
	})
}

func processDocumentXML(xmlContent string, replacements map[string]string) string {
	paragraphs := splitIntoParagraphs(xmlContent)

	for i, paragraph := range paragraphs {
		if strings.Contains(paragraph, "<w:t>") || strings.Contains(paragraph, "<w:t ") {
			plainText := extractTextFromParagraph(paragraph)

			if ContainsAnyKeyword(plainText, replacements) {
				positionMap := buildPositionMap(paragraph)
				paragraphs[i] = ApplyReplacements(paragraph, plainText, replacements, positionMap)
			}
		}
	}
	return strings.Join(paragraphs, "")
}

func splitIntoParagraphs(xmlContent string) []string {
	re := regexp.MustCompile(`(<w:p\b[^>]*>(?:.*?)</w:p>)`)
	var result []string
	lastEnd := 0

	for _, p := range re.FindAllStringIndex(xmlContent, -1) {
		if p[0] > lastEnd {
			result = append(result, xmlContent[lastEnd:p[0]])
		}

		result = append(result, xmlContent[p[0]:p[1]])
		lastEnd = p[1]
	}

	if lastEnd < len(xmlContent) {
		result = append(result, xmlContent[lastEnd:])
	}

	return result
}

func extractTextFromParagraph(paragraph string) string {
	re := regexp.MustCompile(`<w:t(?:\s[^>]*)?>(.*?)</w:t>`)
	matches := re.FindAllStringSubmatch(paragraph, -1)

	var text strings.Builder
	for _, match := range matches {
		if len(match) > 1 {
			text.WriteString(match[1])
		}
	}

	return text.String()
}

func buildPositionMap(paragraph string) map[int]int {
	positionMap := make(map[int]int)

	re := regexp.MustCompile(`<w:t(?:\s[^>]*)?>(.*?)</w:t>`)
	matches := re.FindAllStringSubmatchIndex(paragraph, -1)

	plainPos := 0
	for _, match := range matches {
		if len(match) >= 4 {
			contentStart := match[2] // Start of content group
			contentEnd := match[3]   // End of content group

			textContent := paragraph[contentStart:contentEnd]
			for i := 0; i < len(textContent); i++ {
				positionMap[plainPos+i] = contentStart + i
			}
			plainPos += len(textContent)
		}
	}
	return positionMap
}
