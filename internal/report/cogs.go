package report

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
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

	inventoryShortFall, err := g.inventoryRepo.ComputeInventoryShortFall()
	if err != nil {
		return fmt.Errorf("error computing inventory short fall report. error: %w", err)
	}

	//Compute material model count
	materialModelCount, err := g.inventoryRepo.ComputeMaterialModelCount()
	if err != nil {
		return fmt.Errorf("error computing material model count: %w", err)
	}

	// Use priceData, tseMapping, and creditData in your COGS calculation logic here
	reportFile := excel.NewFile()

	outputDir := utils.GenerateOutputPath(g.cfg.OutputDir, "inventory_report")
	if err := g.writeInventoryReport(reportFile, outputDir, inventoryShortFall); err != nil {
		return fmt.Errorf("error writing inventory report: %w", err)
	}

	if err := g.writeMaterialModelCountReport(reportFile, outputDir, materialModelCount); err != nil {
		return fmt.Errorf("error writing inventory report: %w", err)
	}
	fmt.Printf("COGS report generated successfully: %s\n", outputDir)
	return nil
}

func (g *COGSReportGenerator) writeInventoryReport(f *excelize.File, outputDir string, inventoryShortFallData map[string]*repository.InventoryShortFallRepo) error {
	inventoryShortFallSheet := "Inventory ShortFall"
	// Create a new sheet
	if _, err := f.NewSheet(inventoryShortFallSheet); err != nil {
		return fmt.Errorf("error creating new sheet: %w", err)
	}

	f.DeleteSheet("Sheet1")
	if err := os.MkdirAll(outputDir, os.ModePerm); err != nil {
		return fmt.Errorf("error creating output directory: %w", err)
	}

	headers := []string{"Dealer Code", "Dealer Name", "TSE", "Total Inventory Cost(₹)", "Total Credit Due(₹)", "Inventory Shortfall (₹)"}
	if err := excel.WriteHeaders(f, inventoryShortFallSheet, headers); err != nil {
		return err
	}
	// Custom number format for Indian numbering
	inrFormat := "#,##,##0"
	numberStyle, _ := f.NewStyle(&excelize.Style{
		CustomNumFmt: &inrFormat, // Custom number format for Indian numbering
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})

	// Create a new redStyle
	redStyle, _ := f.NewStyle(&excelize.Style{
		Fill:         excelize.Fill{Type: "pattern", Color: []string{"FF9999"}, Pattern: 1}, // Light red background
		CustomNumFmt: &inrFormat,                                                            // Custom number format for Indian numbering
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
		Font: &excelize.Font{
			Bold: true, // Set text to bold
		},
	})

	// Convert map to slice for sorting
	inventorySlice := make([]*repository.InventoryShortFallRepo, 0, len(inventoryShortFallData))
	for _, data := range inventoryShortFallData {
		inventorySlice = append(inventorySlice, data)
	}

	// Sort by TSE and then by Cost-Credit Difference
	sort.Slice(inventorySlice, func(i, j int) bool {
		if inventorySlice[i].TSE == inventorySlice[j].TSE {
			return inventorySlice[i].InventoryShortfall < inventorySlice[j].InventoryShortfall
		}
		return inventorySlice[i].TSE < inventorySlice[j].TSE
	})

	row := 2
	for _, data := range inventorySlice {
		cellData := []interface{}{
			data.DealerCode,
			data.DealerName,
			data.TSE,
			data.TotalInventoryCost,
			data.TotalCreditDue,
			data.InventoryShortfall,
		}
		if err := excel.WriteRow(f, inventoryShortFallSheet, row, cellData); err != nil {
			return err
		}

		// Apply number style to numeric columns (0-7 Days to Total Credit)
		for col := 3; col <= 5; col++ { // Columns C (3) to I (5)
			cell := fmt.Sprintf("%s%d", string('A'+col), row) // Convert column index to letter
			var style int
			if data.InventoryShortfall < 0 { // Check if InventoryShortfall is negative
				style = redStyle // Use redStyle for negative values
			} else {
				style = numberStyle // Use numberStyle for non-negative values
			}
			if err := f.SetCellStyle(inventoryShortFallSheet, cell, cell, style); err != nil {
				return fmt.Errorf("error setting style for cell %s: %w", cell, err)
			}
		}
		row++
	}
	excel.AdjustColumnWidths(f, inventoryShortFallSheet)
	fileName := "inventory_report.xlsx"
	outputPath := filepath.Join(outputDir, fileName)
	return f.SaveAs(outputPath)
}

func (g *COGSReportGenerator) writeMaterialModelCountReport(f *excelize.File, outputDir string, materialModelCount map[string]*repository.ModelCountRepo) error {
	sheetName := "Material Model Count"
	if _, err := f.NewSheet(sheetName); err != nil {
		return fmt.Errorf("error creating new sheet for material model count: %w", err)
	}
	f.DeleteSheet("Sheet1")

	headers := []string{"Dealer Code", "Dealer Name", "Material Code", "SPU Name", "Color", "SKU Spec", "Product Type", "Count"}
	if err := excel.WriteHeaders(f, sheetName, headers); err != nil {
		return err
	}

	// Convert map to slice for sorting
	materialSlice := make([]*repository.ModelCountRepo, 0, len(materialModelCount))
	for _, data := range materialModelCount {
		materialSlice = append(materialSlice, data)
	}

	// Sort by Count in descending order
	sort.Slice(materialSlice, func(i, j int) bool {
		return materialSlice[i].Count > materialSlice[j].Count
	})

	// Custom number format for Indian numbering
	numberFormat := "0"
	numberStyle, _ := f.NewStyle(&excelize.Style{
		CustomNumFmt: &numberFormat, // Custom number format for Indian numbering
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})

	row := 2
	for _, data := range materialSlice {
		cellData := []interface{}{
			data.DealerCode,
			data.DealerName,
			data.MaterialCode,
			data.SPUName,
			data.Color,
			data.SKUSpec,
			data.ProductType,
			data.Count, // Count
		}
		if err := excel.WriteRow(f, sheetName, row, cellData); err != nil {
			return err
		}

		// Apply number style to numeric columns (only for Material Code)
		cell := fmt.Sprintf("%s%d", "C", row) // Column C for Material Code
		if err := f.SetCellStyle(sheetName, cell, cell, numberStyle); err != nil {
			return fmt.Errorf("error setting style for cell %s: %w", cell, err)
		}

		row++
	}

	excel.AdjustColumnWidths(f, sheetName)
	fileName := "inventory_report.xlsx"
	outputPath := filepath.Join(outputDir, fileName)
	return f.SaveAs(outputPath)
}
