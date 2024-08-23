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

	generator, err := report.NewReportGenerator("growth", cfg)
	if err != nil {
		log.Fatalf("Failed to create Growth report generator: %v", err)
	}

	if err := generator.Generate(); err != nil {
		log.Fatalf("Failed to generate Growth report: %v", err)
	}

	log.Println("Growth report generated successfully")
}
