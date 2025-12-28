package internal

import (
	"archive/zip"
	"regexp"
	"strings"
)

// ProcessPptxZipFile processes a single file from a PPTX zip archive.
func ProcessPptxZipFile(file *zip.File, zipWriter *zip.Writer, replacements map[string]string) error {
	return ProcessZipFile(file, zipWriter, func(fileName string, content []byte) []byte {
		// Process slide files - each slide is a separate XML file
		if strings.HasPrefix(fileName, "ppt/slides/slide") && strings.HasSuffix(fileName, ".xml") {
			processedContent := processSlideXML(string(content), replacements)
			return []byte(processedContent)
		}
		// For other files, copy as-is
		return content
	})
}

func processSlideXML(xmlContent string, replacements map[string]string) string {
	textFrames := splitIntoTextFrames(xmlContent)

	for i, frame := range textFrames {
		if strings.Contains(frame, "<a:t>") || strings.Contains(frame, "<a:t ") {
			plainText := extractTextFromFrame(frame)

			if ContainsAnyKeyword(plainText, replacements) {
				positionMap := buildFramePositionMap(frame)
				textFrames[i] = ApplyReplacements(frame, plainText, replacements, positionMap)
			}
		}
	}
	return strings.Join(textFrames, "")
}

func splitIntoTextFrames(xmlContent string) []string {
	// Match <a:p> elements (paragraphs) in DrawingML namespace
	// Text in PPTX is organized in paragraphs within text bodies
	re := regexp.MustCompile(`(<a:p\b[^>]*>(?:.*?)</a:p>)`)
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

func extractTextFromFrame(frame string) string {
	// Extract text from <a:t> tags (DrawingML text runs)
	// Note: PPTX uses DrawingML namespace with 'a:' prefix
	re := regexp.MustCompile(`<a:t(?:\s[^>]*)?>(.*?)</a:t>`)
	matches := re.FindAllStringSubmatch(frame, -1)

	var text strings.Builder
	for _, match := range matches {
		if len(match) > 1 {
			text.WriteString(match[1])
		}
	}

	return text.String()
}

func buildFramePositionMap(frame string) map[int]int {
	positionMap := make(map[int]int)

	// Build position map for <a:t> tags in DrawingML paragraphs
	re := regexp.MustCompile(`<a:t(?:\s[^>]*)?>(.*?)</a:t>`)
	matches := re.FindAllStringSubmatchIndex(frame, -1)

	plainPos := 0
	for _, match := range matches {
		if len(match) >= 4 {
			contentStart := match[2] // Start of content group
			contentEnd := match[3]   // End of content group

			textContent := frame[contentStart:contentEnd]
			for i := 0; i < len(textContent); i++ {
				positionMap[plainPos+i] = contentStart + i
			}
			plainPos += len(textContent)
		}
	}
	return positionMap
}
