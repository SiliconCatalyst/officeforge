package benchmarks

import (
	"fmt"
	"os"
	"testing"
	"time"

	"github.com/siliconcatalyst/officeforge/docx"
)

func TestRealWorldPerformance(t *testing.T) {
	os.MkdirAll("testdata/output", 0755)
	defer os.RemoveAll("testdata/output")

	// Your actual template keywords
	replacements := map[string]string{
		"{{CONTRACT_DATE}}":    "2024-12-30",
		"{{CLIENT_NAME}}":      "John Doe",
		"{{CLIENT_COMPANY}}":   "Acme Corporation",
		"{{CLIENT_EMAIL}}":     "john.doe@acme.com",
		"{{CLIENT_PHONE}}":     "+1-555-0123",
		"{{PROJECT_NAME}}":     "Website Redesign",
		"{{CONTRACT_AMOUNT}}":  "$50,000.00",
		"{{PROJECT_DEADLINE}}": "2025-03-31",
		"{{PAYMENT_TERMS}}":    "Net 30 days",
	}

	iterations := 1000

	start := time.Now()
	for i := 0; i < iterations; i++ {
		err := docx.ProcessDocxMulti(
			"../testdata/template.docx",
			fmt.Sprintf("testdata/output/contract_%d.docx", i),
			replacements,
		)
		if err != nil {
			t.Fatalf("Failed at iteration %d: %v", i, err)
		}
	}
	duration := time.Since(start)

	totalReplacements := len(replacements) * iterations
	docsPerSecond := float64(iterations) / duration.Seconds()
	replacementsPerSecond := float64(totalReplacements) / duration.Seconds()

	fmt.Printf("\n╔══════════════════════════════════════════╗\n")
	fmt.Printf("║      PERFORMANCE TEST RESULTS            ║\n")
	fmt.Printf("╠══════════════════════════════════════════╣\n")
	fmt.Printf("║ Documents processed:   %6d           ║\n", iterations)
	fmt.Printf("║ Keywords per doc:      %6d           ║\n", len(replacements))
	fmt.Printf("║ Total replacements:    %6d           ║\n", totalReplacements)
	fmt.Printf("║ Time taken:            %6.2fs          ║\n", duration.Seconds())
	fmt.Printf("╠══════════════════════════════════════════╣\n")
	fmt.Printf("║ Throughput:            %6.0f docs/s    ║\n", docsPerSecond)
	fmt.Printf("║ Replacements/sec:      %6.0f/s         ║\n", replacementsPerSecond)
	fmt.Printf("║ Avg per doc:           %6.2fms         ║\n", duration.Seconds()*1000/float64(iterations))
	fmt.Printf("╚══════════════════════════════════════════╝\n\n")

	// Assert minimum performance
	if docsPerSecond < 100 {
		t.Errorf("Performance regression: expected >100 docs/sec, got %.0f", docsPerSecond)
	}
}

// Benchmark for continuous monitoring
func BenchmarkDocxContract(b *testing.B) {
	replacements := map[string]string{
		"{{CONTRACT_DATE}}":    "2024-12-30",
		"{{CLIENT_NAME}}":      "John Doe",
		"{{CLIENT_COMPANY}}":   "Acme Corporation",
		"{{CLIENT_EMAIL}}":     "john.doe@acme.com",
		"{{CLIENT_PHONE}}":     "+1-555-0123",
		"{{PROJECT_NAME}}":     "Website Redesign",
		"{{CONTRACT_AMOUNT}}":  "$50,000.00",
		"{{PROJECT_DEADLINE}}": "2025-03-31",
		"{{PAYMENT_TERMS}}":    "Net 30 days",
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		docx.ProcessDocxMulti(
			"../testdata/template.docx",
			fmt.Sprintf("testdata/output/bench_%d.docx", i),
			replacements,
		)
	}
}
