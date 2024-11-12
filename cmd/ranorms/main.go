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

	generator, err := report.NewReportGenerator("ranorms", cfg)
	if err != nil {
		log.Fatalf("Failed to create RA Norms report generator: %v", err)
	}

	if err := generator.Generate(); err != nil {
		log.Fatalf("Failed to generate RA Norms report: %v", err)
	}

	log.Println("RA Norms report generated successfully")
	time.Sleep(5 * time.Second) // {{ edit_2 }}

}
