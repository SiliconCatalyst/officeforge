package internal

import (
	"archive/zip"
	"io"
	"log"
	"sort"
	"strings"
)

type replacementPoint struct {
	startPos    int
	endPos      int
	replacement string
}

// ProcessZipFile processes a single file from a zip archive and writes it to the output zip.
// fileProcessor is a function that processes the content if needed, otherwise returns it unchanged.
func ProcessZipFile(file *zip.File, zipWriter *zip.Writer, fileProcessor func(string, []byte) []byte) error {
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

	// Process the content (or pass through unchanged)
	processedContent := fileProcessor(file.Name, content)
	_, err = writer.Write(processedContent)

	return err
}

// FindReplacementPoints identifies all positions in the text where keywords should be replaced.
func FindReplacementPoints(text string, replacements map[string]string) []replacementPoint {
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

// ApplyReplacements applies replacement points to a paragraph/element using a position map.
func ApplyReplacements(element string, plainText string, replacements map[string]string, positionMap map[int]int) string {
	replacementPoints := FindReplacementPoints(plainText, replacements)
	elementLength := len(element)

	// Sort replacement points by position (descending) to avoid position shifts
	sort.Slice(replacementPoints, func(i, j int) bool {
		return replacementPoints[i].startPos > replacementPoints[j].startPos
	})

	for _, rp := range replacementPoints {
		xmlStartPos, xmlEndPos := FindXMLPositions(rp.startPos, rp.endPos, positionMap)

		if xmlStartPos < 0 || xmlEndPos <= xmlStartPos {
			log.Printf("Skipping invalid replacement: start=%d, end=%d", xmlStartPos, xmlEndPos)
			continue
		}

		if xmlStartPos >= elementLength {
			log.Printf("Start position %d out of bounds (length: %d)", xmlStartPos, elementLength)
			continue
		}

		if xmlEndPos > elementLength {
			log.Printf("Adjusting end position from %d to %d (element length)", xmlEndPos, elementLength)
			xmlEndPos = elementLength
		}

		// Double-check to avoid slice bounds error
		if xmlStartPos <= xmlEndPos && xmlStartPos < elementLength && xmlEndPos <= elementLength {
			element = element[:xmlStartPos] + rp.replacement + element[xmlEndPos:]
		} else {
			log.Printf("Skipping replacement due to invalid positions: start=%d, end=%d, length=%d",
				xmlStartPos, xmlEndPos, elementLength)
		}
	}
	return element
}

// FindXMLPositions converts plain text positions to XML positions using a position map.
func FindXMLPositions(plainStartPos, plainEndPos int, positionMap map[int]int) (int, int) {
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

// ContainsAnyKeyword checks if text contains any of the keywords from the replacements map.
func ContainsAnyKeyword(text string, replacements map[string]string) bool {
	for keyword := range replacements {
		if strings.Contains(text, keyword) {
			return true
		}
	}
	return false
}
