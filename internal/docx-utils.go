package internal

import (
	"archive/zip"
	"io"
	"log"
	"regexp"
	"sort"
	"strings"
)

type replacementPoint struct {
	startPos    int
	endPos      int
	replacement string
}

func ProcessZipFile(file *zip.File, zipWriter *zip.Writer, replacements map[string]string) error {
	// Open the file from the zip
	rc, err := file.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	// Read the content
	content, err := io.ReadAll(rc)
	if err != nil {
		return err
	}

	// Create corresponding file in output zip
	writer, err := zipWriter.Create(file.Name)
	if err != nil {
		return err
	}

	// If this is the main document XML file, process it for replacements
	if file.Name == "word/document.xml" {
		processedContent := processDocumentXML(string(content), replacements)
		_, err = writer.Write([]byte(processedContent))
	} else {
		// For other files, copy as-is
		_, err = writer.Write(content)
	}

	return err
}

func processDocumentXML(xmlContent string, replacements map[string]string) string {
	paragraphs := splitIntoParagraphs(xmlContent)

	for i, paragraph := range paragraphs {
		if strings.Contains(paragraph, "<w:t>") || strings.Contains(paragraph, "<w:t ") {
			plainText := extractTextFromParagraph(paragraph)
			needsReplacement := false
			for keyword := range replacements {
				if strings.Contains(plainText, keyword) {
					needsReplacement = true
					break
				}
			}

			if needsReplacement {
				paragraphs[i] = replaceParagraphText(paragraph, plainText, replacements)
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

func replaceParagraphText(paragraph, plainText string, replacements map[string]string) string {
	positionMap := buildPositionMap(paragraph)
	replacementPoints := findReplacementPoints(plainText, replacements)
	paragraphLength := len(paragraph)

	// Sort replacement points by position (descending) to avoid position shifts
	sort.Slice(replacementPoints, func(i, j int) bool {
		return replacementPoints[i].startPos > replacementPoints[j].startPos
	})

	for _, rp := range replacementPoints {
		xmlStartPos, xmlEndPos := findXMLPositions(rp.startPos, rp.endPos, positionMap)

		if xmlStartPos < 0 || xmlEndPos <= xmlStartPos {
			log.Printf("Skipping invalid replacement: start=%d, end=%d", xmlStartPos, xmlEndPos)
			continue
		}

		if xmlStartPos >= paragraphLength {
			log.Printf("Start position %d out of bounds (length: %d)", xmlStartPos, paragraphLength)
			continue
		}

		if xmlEndPos > paragraphLength {
			log.Printf("Adjusting end position from %d to %d (paragraph length)", xmlEndPos, paragraphLength)
			xmlEndPos = paragraphLength
		}

		// Double-check to avoid slice bounds error
		if xmlStartPos <= xmlEndPos && xmlStartPos < paragraphLength && xmlEndPos <= paragraphLength {
			paragraph = paragraph[:xmlStartPos] + rp.replacement + paragraph[xmlEndPos:]
		} else {
			log.Printf("Skipping replacement due to invalid positions: start=%d, end=%d, length=%d",
				xmlStartPos, xmlEndPos, paragraphLength)
		}
	}
	return paragraph
}

func findReplacementPoints(text string, replacements map[string]string) []replacementPoint {
	var points []replacementPoint
	for keyword, replacement := range replacements {
		index := 0
		for {
			pos := strings.Index(text[index:], keyword)
			if pos == -1 {
				break
			}

			startPos := index + pos
			endPos := startPos + len(keyword)

			points = append(points, replacementPoint{
				startPos:    startPos,
				endPos:      endPos,
				replacement: replacement,
			})
			index = endPos
		}
	}
	return points
}

func findXMLPositions(plainStartPos, plainEndPos int, positionMap map[int]int) (int, int) {
	// Find the actual positions in the map
	xmlStartPos := -1
	xmlEndPos := -1

	// Find the exact start position or nearest position
	for i := plainStartPos; i < plainEndPos; i++ {
		if pos, exists := positionMap[i]; exists {
			xmlStartPos = pos
			break
		}
	}

	if xmlStartPos == -1 {
		log.Printf("No start position found for plaintext range %d:%d", plainStartPos, plainEndPos)
		return -1, -1
	}

	// Find the exact end position or nearest position before the end
	for i := plainEndPos - 1; i >= plainStartPos; i-- {
		if pos, exists := positionMap[i]; exists {
			xmlEndPos = pos + 1 // +1 because end position is exclusive
			break
		}
	}

	// If no end position found, use start position + 1 as fallback
	if xmlEndPos == -1 {
		log.Printf("No end position found, using start+1 as fallback")
		xmlEndPos = xmlStartPos + 1
	}

	// Final safety check
	if xmlEndPos <= xmlStartPos {
		log.Printf("XML positions invalid: start=%d, end=%d, adjusting end", xmlStartPos, xmlEndPos)
		xmlEndPos = xmlStartPos + 1
	}

	return xmlStartPos, xmlEndPos
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
