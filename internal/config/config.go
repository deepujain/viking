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
	SalesReport     string
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
			DealerInfo: filepath.Join(dataDir, "Retailer Metadata.xlsx"),
			TSEMapping: filepath.Join(dataDir, "Retailer Metadata.xlsx"),
			PriceList:  filepath.Join(dataDir, "ProductPriceList.xlsx"),
		},
		ReportFiles: ReportFiles{
			CreditReport: CreditReportFiles{
				Bills: filepath.Join(dataDir, "Bills.xlsx"),
			},
			GrowthReport: GrowthReportFiles{
				MTDSO:  filepath.Join(dataDir, "MTD-SO.xlsx"),
				LMTDSO: filepath.Join(dataDir, "LMTD-SO.xlsx"),
				MTDST:  filepath.Join(dataDir, "MTD-ST.xlsx"),
				LMTDST: filepath.Join(dataDir, "LMTD-ST.xlsx"),
			},
			InventoryReport: filepath.Join(dataDir, "DealerInventory.xlsx"),
			PriceListFile:   filepath.Join(dataDir, "ZD PRICE LIST.xlsx"),
			SalesReport:     filepath.Join(dataDir, "Sales.xlsx"),
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
