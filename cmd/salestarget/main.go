package main

import (
	"log"
	"time"

	"viking-reports/internal/config"
	"viking-reports/internal/report"
)

func main() {
	cfg, err := config.Load()
	if err != nil {
		log.Fatalf("Failed to load configuration: %v", err)
	}

	generator, err := report.NewReportGenerator("salestarget", cfg) // Updated report type
	if err != nil {
		log.Fatalf("Failed to create Sales Target report generator: %v", err) // Updated error message
	}

	if err := generator.Generate(); err != nil {
		log.Fatalf("Failed to generate Sales Target report: %v", err) // Updated error message
	}
	time.Sleep(5 * time.Second) // Unchanged
}
