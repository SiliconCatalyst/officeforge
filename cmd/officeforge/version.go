package main

import "fmt"

var (
	Version   = "v0.0.0"
	BuildDate = "unknown"
	Commit    = "none"
)

func printVersion() {
	// Outputs: OfficeForge v0.5.1 (29 - December - 2025) [commit: a1b2c3d]
	fmt.Printf("OfficeForge %s (%s) [commit: %s]\n", Version, BuildDate, Commit)
}
