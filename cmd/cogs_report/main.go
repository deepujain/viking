package main

import (
	"log"

	"viking-reports/internal/config"
	"viking-reports/internal/report"
)

func main() {
	cfg, err := config.Load()
	if err != nil {
		log.Fatalf("Failed to load configuration: %v", err)
	}

	generator, err := report.NewReportGenerator("cogs", cfg)
	if err != nil {
		log.Fatalf("Failed to create COGS report generator: %v", err)
	}

	if err := generator.Generate(); err != nil {
		log.Fatalf("Failed to generate COGS report: %v", err)
	}

}
