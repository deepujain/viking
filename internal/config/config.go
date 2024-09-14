package config

import (
	"os"
	"path/filepath"
)

// Config holds the application configuration
type Config struct {
	DataDir     string
	OutputDir   string
	CommonFiles CommonFiles
	ReportFiles ReportFiles
}

// CommonFiles holds paths to common files used across reports
type CommonFiles struct {
	DealerInfo string
	TSEMapping string
	PriceList  string
}

// ReportFiles holds paths to report-specific files
type ReportFiles struct {
	CreditReport    CreditReportFiles
	GrowthReport    GrowthReportFiles
	InventoryReport string
	PriceListFile   string
}

// CreditReportFiles holds paths to credit report files
type CreditReportFiles struct {
	Bills string
}

// GrowthReportFiles holds paths to growth report files
type GrowthReportFiles struct {
	MTDSO  string
	LMTDSO string
	MTDST  string
	LMTDST string
}

// Load returns a new Config struct with default values
func Load() (*Config, error) {
	dataDir := filepath.Join("..", "data")
	outputDir := "."

	config := &Config{
		DataDir:   dataDir,
		OutputDir: outputDir,
		CommonFiles: CommonFiles{
			DealerInfo: filepath.Join(dataDir, "common", "Retailer Metadata.xlsx"),
			TSEMapping: filepath.Join(dataDir, "common", "Retailer Metadata.xlsx"),
			PriceList:  filepath.Join(dataDir, "common", "ProductPriceList.xlsx"),
		},
		ReportFiles: ReportFiles{
			CreditReport: CreditReportFiles{
				Bills: filepath.Join(dataDir, "credit_report", "Bills.xlsx"),
			},
			GrowthReport: GrowthReportFiles{
				MTDSO:  filepath.Join(dataDir, "growth_report", "MTD-SO.xlsx"),
				LMTDSO: filepath.Join(dataDir, "growth_report", "LMTD-SO.xlsx"),
				MTDST:  filepath.Join(dataDir, "growth_report", "MTD-ST.xlsx"),
				LMTDST: filepath.Join(dataDir, "growth_report", "LMTD-ST.xlsx"),
			},
			InventoryReport: filepath.Join(dataDir, "cogs_report", "DealerInventory.xlsx"),
			PriceListFile:   filepath.Join(dataDir, "price_list", "PRICE LIST AS ON 02.09.2024.xlsx"),
		},
	}

	// Ensure directories exist
	if err := os.MkdirAll(config.DataDir, os.ModePerm); err != nil {
		return nil, err
	}
	if err := os.MkdirAll(config.OutputDir, os.ModePerm); err != nil {
		return nil, err
	}

	return config, nil
}
