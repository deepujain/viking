package report

import (
	"fmt"
	"os"
	"path/filepath"
	"viking-reports/internal/config"
	"viking-reports/internal/repository"
	"viking-reports/internal/utils"
	"viking-reports/pkg/excel"

	"github.com/xuri/excelize/v2"
)

type COGSReportGenerator struct {
	cfg              *config.Config
	inventoryRepo    repository.InventoryRepository
	creditRepo       repository.CreditRepository
	tseMappingRepo   repository.TSEMappingRepository
	productPriceRepo repository.ProductPriceRepository
}

func NewCOGSReportGenerator(cfg *config.Config) *COGSReportGenerator {
	priceRepo := repository.NewExcelProductPriceRepository(cfg.CommonFiles.PriceList)
	tseMappingRepo := repository.NewExcelTSEMappingRepository(cfg.CommonFiles.TSEMapping)

	priceData, _ := priceRepo.GetProductPrices()
	tseMapping, _ := tseMappingRepo.GetRetailerCodeToTSEMap()

	return &COGSReportGenerator{
		cfg:              cfg,
		inventoryRepo:    repository.NewExcelInventoryRepository(cfg.ReportFiles.InventoryReport, priceData, tseMapping),
		creditRepo:       repository.NewExcelCreditRepository(cfg.DataDir),
		tseMappingRepo:   tseMappingRepo,
		productPriceRepo: priceRepo,
	}
}

func (g *COGSReportGenerator) Generate() error {
	fmt.Println("Generating COGS report...")

	inventoryData, err := g.inventoryRepo.GetInventoryData()
	if err != nil {
		return fmt.Errorf("error reading inventory data: %w", err)
	}

	// Use priceData, tseMapping, and creditData in your COGS calculation logic here

	outputDir := utils.GenerateOutputPath(g.cfg.OutputDir, "inventory_cost_report")
	if err := g.writeInventoryReport(outputDir, inventoryData); err != nil {
		return fmt.Errorf("error writing inventory report: %w", err)
	}

	fmt.Printf("COGS report generated successfully: %s\n", outputDir)
	return nil
}

func (g *COGSReportGenerator) writeInventoryReport(outputDir string, inventoryData map[string]*repository.InventoryData) error {
	f := excel.NewFile()
	sheetName := "Inventory Report"
	// Create a new sheet
	if _, err := f.NewSheet(sheetName); err != nil {
		return fmt.Errorf("error creating new sheet: %w", err)
	}
	f.DeleteSheet("Sheet1")
	if err := os.MkdirAll(outputDir, os.ModePerm); err != nil {
		return fmt.Errorf("error creating output directory: %w", err)
	}

	headers := []string{"Dealer Code", "Dealer Name", "TSE", "Total Inventory Cost(₹)", "Total Credit Due(₹)", "Cost-Credit Difference(₹)"}
	if err := excel.WriteHeaders(f, sheetName, headers); err != nil {
		return err
	}
	// Custom number format for Indian numbering
	inrFormat := "#,##,##0.00"
	numberStyle, _ := f.NewStyle(&excelize.Style{
		CustomNumFmt: &inrFormat, // Custom number format for Indian numbering
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	row := 2
	for _, data := range inventoryData {
		cellData := []interface{}{
			data.DealerCode,
			data.DealerName,
			data.TSE,
			data.TotalInventoryCost,
			data.TotalCreditDue,
			data.CostCreditDifference,
		}
		if err := excel.WriteRow(f, sheetName, row, cellData); err != nil {
			return err
		}

		// Apply number style to numeric columns (0-7 Days to Total Credit)
		for col := 3; col <= 5; col++ { // Columns C (3) to I (5)
			cell := fmt.Sprintf("%s%d", string('A'+col), row) // Convert column index to letter
			if err := f.SetCellStyle(sheetName, cell, cell, numberStyle); err != nil {
				return fmt.Errorf("error setting style for cell %s: %w", cell, err)
			}
		}
		row++
	}
	excel.AdjustColumnWidths(f, sheetName)
	fileName := "daily_inventory_cost_report.xlsx"
	outputPath := filepath.Join(outputDir, fileName)
	return f.SaveAs(outputPath)
}
