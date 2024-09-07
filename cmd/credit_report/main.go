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

	generator, err := report.NewReportGenerator("credit", cfg)
	if err != nil {
		log.Fatalf("Failed to create Credit report generator: %v", err)
	}

	if err := generator.Generate(); err != nil {
		log.Fatalf("Failed to generate Credit report: %v", err)
	}
	time.Sleep(5 * time.Second) // {{ edit_2 }}
}
