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

type ZSOReportGenerator struct {
	cfg            *config.Config
	inventoryRepo  repository.InventoryRepository
	salesRepo      repository.SalesRepository
	tseMappingRepo repository.TSEMappingRepository
}

func NewZSOReportGenerator(cfg *config.Config) *ZSOReportGenerator {

	return &ZSOReportGenerator{
		cfg:            cfg,
		inventoryRepo:  repository.NewSPUInventoryRepository(cfg.ReportFiles.InventoryReport),
		salesRepo:      repository.NewExcelSalesRepository(), // Use one of the file paths
		tseMappingRepo: repository.NewExcelTSEMappingRepository(cfg.CommonFiles.TSEMapping),
	}
}

func (g *ZSOReportGenerator) Generate() error {
	fmt.Println("Generating ZSO report...")

	// Fetch inventory and sales data
	fmt.Print("Input: Fetching per dealer per SPU current inventory count")
	dealerSPUInventory, err := g.inventoryRepo.ComputeDealerSPUInventory()
	if err != nil {
		return fmt.Errorf("error reading inventory data: %w", err)
	}

	fmt.Print("Input: Fetching per dealer per SPU last month to date sell out (SO) count")
	lmtdDealerSPUSales, err := g.salesRepo.GetDealerSPUSales(g.cfg.ReportFiles.GrowthReport.LMTDSO)
	if err != nil {
		return fmt.Errorf("error reading LMTD sales data: %w", err)
	}

	fmt.Print("Input: Fetching per dealer per SPU month to date sell out (ST) count")
	mtdDealerSPUSales, err := g.salesRepo.GetDealerSPUSales(g.cfg.ReportFiles.GrowthReport.MTDSO)
	if err != nil {
		return fmt.Errorf("error reading MTD sales data: %w", err)
	}

	tseMapping, err := g.tseMappingRepo.GetRetailerNameToTSEMap()
	if err != nil {
		return fmt.Errorf("error reading TSE mapping: %w", err)
	}
	// Combine LMTD and MTD sales
	allSales := append(mapToSlice(lmtdDealerSPUSales), mapToSlice(mtdDealerSPUSales)...)

	fmt.Println("Start computation of zso report.")
	// Build inventory map
	inventoryMap := make(map[string]int)
	for _, inventory := range dealerSPUInventory {
		dealerSPUKey := inventory.DealerName + inventory.SPUName
		inventoryMap[dealerSPUKey] = inventory.Count
	}

	fmt.Println("Identifying zso.")
	// Track ZSO data and relevant model names
	zsoData := make(map[string]map[string]string)
	zsoModelNames := make(map[string]struct{})

	for _, salesData := range allSales {
		dealer := salesData.DealerName
		model := salesData.SPUName

		if _, dealerExists := zsoData[dealer]; !dealerExists {
			zsoData[dealer] = make(map[string]string)
		}

		// Only mark models with zero inventory as ZSO
		dealerSPUKey := dealer + model
		if inventoryMap[dealerSPUKey] == 0 {
			zsoData[dealer][model] = "250" // Mark ZSO
			zsoModelNames[model] = struct{}{}
		}
	}

	// Generate and save the Excel report
	outputDir := utils.GenerateOutputPath(g.cfg.OutputDir, "zso_report")
	err = g.writeZSOReport(zsoData, zsoModelNames, tseMapping, outputDir)
	if err != nil {
		return fmt.Errorf("error writing ZSO report: %w", err)
	}

	return nil
}

// Helper to map sales data to slice format
func mapToSlice(dealerSalesMap map[string]*repository.DealerSPUSales) []repository.DealerSPUSales {
	var dealerSalesSlice []repository.DealerSPUSales
	for _, sales := range dealerSalesMap {
		dealerSalesSlice = append(dealerSalesSlice, *sales)
	}
	return dealerSalesSlice
}

func (g *ZSOReportGenerator) writeZSOReport(zsoData map[string]map[string]string, zsoModelNames map[string]struct{}, tseMapping map[string]string, outputDir string) error {
	f := excelize.NewFile()
	sheetName := "ZSO Report"
	f.NewSheet(sheetName)
	f.DeleteSheet("Sheet1")

	if err := os.MkdirAll(outputDir, os.ModePerm); err != nil {
		return fmt.Errorf("error creating output directory: %w", err)
	}

	// Define styles
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Fill:   excelize.Fill{Type: "pattern", Color: []string{"FFFF00"}, Pattern: 1}, // Yellow background for header
		Font:   &excelize.Font{Bold: true},
		Border: createBorderStyle(),
	})
	zsoCellStyle, _ := f.NewStyle(&excelize.Style{
		Fill:   excelize.Fill{Type: "pattern", Color: []string{"FF9999"}, Pattern: 1}, // Light red for ZSO cells
		Border: createBorderStyle(),
	})
	defaultCellStyle, _ := f.NewStyle(&excelize.Style{
		Border: createBorderStyle(),
	})

	// Build headers with ZSO model names only
	headers := []string{"TSE", "Dealer Name"}
	for model := range zsoModelNames {
		headers = append(headers, model)
	}
	headers = append(headers, "Total ZSO")

	// Write headers with style
	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		f.SetCellValue(sheetName, cell, header)
		f.SetCellStyle(sheetName, cell, cell, headerStyle)
	}
	// Create a slice to hold dealers for sorting
	var dealers []string
	for dealer := range zsoData {
		dealers = append(dealers, dealer)
	}
	// Sort dealers by TSE
	sort.Slice(dealers, func(i, j int) bool {
		return tseMapping[dealers[i]] < tseMapping[dealers[j]]
	})

	// Write ZSO data
	row := 2
	for _, dealer := range dealers {
		models := zsoData[dealer]
		tseCell, _ := excelize.CoordinatesToCellName(1, row)
		f.SetCellValue(sheetName, tseCell, tseMapping[dealer])
		f.SetCellStyle(sheetName, tseCell, tseCell, defaultCellStyle)

		dealerCell, _ := excelize.CoordinatesToCellName(2, row)
		f.SetCellValue(sheetName, dealerCell, dealer)
		f.SetCellStyle(sheetName, dealerCell, dealerCell, defaultCellStyle)

		col := 3
		totalZSO := 0
		for _, model := range headers[2 : len(headers)-1] {
			cell, _ := excelize.CoordinatesToCellName(col, row)
			if models[model] == "250" {
				f.SetCellValue(sheetName, cell, "ZSO")
				f.SetCellStyle(sheetName, cell, cell, zsoCellStyle)
				totalZSO++
			} else {
				f.SetCellStyle(sheetName, cell, cell, defaultCellStyle)
			}
			col++
		}

		totalCell, _ := excelize.CoordinatesToCellName(col, row)
		f.SetCellValue(sheetName, totalCell, totalZSO)
		f.SetCellStyle(sheetName, totalCell, totalCell, defaultCellStyle)
		row++
	}

	outputPath := filepath.Join(outputDir, "zso_report.xlsx")
	excel.AdjustColumnWidths(f, sheetName)
	return f.SaveAs(outputPath)
}

// Helper function to create border style
func createBorderStyle() []excelize.Border {
	return []excelize.Border{
		{Type: "left", Color: "000000", Style: 1},
		{Type: "top", Color: "000000", Style: 1},
		{Type: "bottom", Color: "000000", Style: 1},
		{Type: "right", Color: "000000", Style: 1},
	}
}
