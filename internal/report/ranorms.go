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

type RANormsReportGenerator struct {
	cfg            *config.Config
	inventoryRepo  repository.InventoryRepository
	tseMappingRepo repository.TSEMappingRepository
}

func NewRANormsReportGenerator(cfg *config.Config) *RANormsReportGenerator {

	return &RANormsReportGenerator{
		cfg:            cfg,
		inventoryRepo:  repository.NewSPUInventoryRepository(cfg.ReportFiles.InventoryReport),
		tseMappingRepo: repository.NewExcelTSEMappingRepository(cfg.CommonFiles.TSEMapping),
	}
}

func (g *RANormsReportGenerator) Generate() error {
	fmt.Println("Generating Retailer Agreement (RA) Norms report...")
	// Define the filter list as a map for easier comparison
	modelsOfInterest := map[string]struct{}{
		"C61":        {},
		"C63":        {},
		"C63 5G":     {},
		"C65 5G":     {},
		"13 5G":      {},
		"13+ 5G":     {},
		"13 Pro+ 5G": {},
		"13 Pro 5G":  {},
		"GT 6T":      {},
		"GT6":        {},
	}
	raRetailers, err := g.tseMappingRepo.GetRARetailersMap()
	if err != nil {
		return fmt.Errorf("error reading TSE mapping: %w", err)
	}
	retailerCodeToTSEMap, err := g.tseMappingRepo.GetRetailerCodeToTSEMap()
	if err != nil {
		return fmt.Errorf("error reading TSE mapping: %w", err)
	}

	retailerCodeToNameMap, err := g.tseMappingRepo.GetRetailerCodeToNameMap()
	if err != nil {
		return fmt.Errorf("error reading TSE mapping: %w", err)
	}
	raDealerInventory, err := g.inventoryRepo.ComputeRADealerSPUInventory(modelsOfInterest, raRetailers)
	if err != nil {
		return fmt.Errorf("error reading inventory data: %w", err)
	}

	// Compute RA norms using the raRetailers map
	fmt.Println("Start computation of RA norms report.")
	fmt.Print("Identifying RA norms for ")
	for model := range modelsOfInterest {
		fmt.Print(model)
		fmt.Print(" ")
	}
	fmt.Println()
	raNormsData, err := g.computeRANorms(raRetailers, raDealerInventory, modelsOfInterest)
	if err != nil {
		return fmt.Errorf("error computing RA norms: %w", err)
	}

	// Write the RA norms refill report to Excel
	// Generate and save the Excel report
	outputDir := utils.GenerateOutputPath(g.cfg.OutputDir, "ranorms_report")

	if err := g.writeRANormsReport(raNormsData, retailerCodeToTSEMap, retailerCodeToNameMap, modelsOfInterest, outputDir); err != nil {
		return fmt.Errorf("error writing RA norms report: %w", err)
	}

	return nil
}

func (g *RANormsReportGenerator) computeRANorms(raRetailers map[string]int, raDealerInventory map[string]*repository.SPUInventoryCount, modelsOfInterest map[string]struct{}) (map[string]map[string]int, error) {
	// Initialize the map to store the computed RA norms for each retailer and SPU
	raNormsData := make(map[string]map[string]int)

	for retailer, countOfRA := range raRetailers {
		// Initialize the sub-map for each retailer if not already done
		if _, exists := raNormsData[retailer]; !exists {
			raNormsData[retailer] = make(map[string]int)
		}

		// Loop over each model of interest
		for model := range modelsOfInterest {
			// Construct the composite key for raDealerInventory by combining retailer and model
			// We are assuming that the dealerCode corresponds to a dealerName somewhere, maybe in a separate map
			// If this assumption is wrong, we would need to resolve it.
			// For this example, I will assume that the dealerCode directly matches the key format in raDealerInventory
			compositeKey := retailer + model // Concatenate retailer and model to form the composite key

			// Get the current inventory for the retailer and model
			currentInventory := 0
			if spuInventory, exists := raDealerInventory[compositeKey]; exists {
				// Assuming SPUInventoryCount has a Count field for the inventory value
				if spuInventory != nil {
					currentInventory = spuInventory.Count
				}
			}

			// Calculate the RA norm refill requirement: 3 * storeCount - currentInventory
			requiredInventory := 3*countOfRA - currentInventory
			if requiredInventory < 0 {
				requiredInventory = 0 // No need to refill if inventory meets or exceeds the RA norms
			}

			// Store the computed refill requirement in the raNormsData map
			raNormsData[retailer][model] = requiredInventory
		}
	}

	return raNormsData, nil
}

// Write the RA Norms refill report to Excel
func (g *RANormsReportGenerator) writeRANormsReport(raNormsData map[string]map[string]int, retailerCodeToTSEMap map[string]string, retailerCodeToNameMap map[string]string, modelsOfInterest map[string]struct{}, outputDir string) error {
	f := excelize.NewFile()
	sheetName := "RA Norms Report"
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
	refillCellStyle, _ := f.NewStyle(&excelize.Style{
		Fill:   excelize.Fill{Type: "pattern", Color: []string{"FFCCCC"}, Pattern: 1}, // Light red for refill cells
		Border: createBorderStyle(),
	})
	defaultCellStyle, _ := f.NewStyle(&excelize.Style{
		Border: createBorderStyle(),
	})

	// Build headers
	headers := []string{"TSE", "Dealer Name"}
	for model := range modelsOfInterest {
		headers = append(headers, model) // Add model names as headers
	}
	headers = append(headers, "Total Refill") // Add total refill column
	// Sort headers alphabetically
	sort.Strings(headers[2:]) // Sort the ZSO models in headers

	// Write headers with style
	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		f.SetCellValue(sheetName, cell, header)
		f.SetCellStyle(sheetName, cell, cell, headerStyle)
	}

	// Sort dealers by TSE name for organized reporting
	var dealers []string
	for dealer := range raNormsData {
		dealers = append(dealers, dealer)
	}
	sort.Slice(dealers, func(i, j int) bool {
		return retailerCodeToTSEMap[dealers[i]] < retailerCodeToTSEMap[dealers[j]]
	})

	// Write RA norms refill data
	row := 2
	for _, dealerCode := range dealers {
		modelRefill := raNormsData[dealerCode]
		tseCell, _ := excelize.CoordinatesToCellName(1, row)
		f.SetCellValue(sheetName, tseCell, retailerCodeToTSEMap[dealerCode])
		f.SetCellStyle(sheetName, tseCell, tseCell, defaultCellStyle)

		dealerCell, _ := excelize.CoordinatesToCellName(2, row)
		f.SetCellValue(sheetName, dealerCell, retailerCodeToNameMap[dealerCode])
		f.SetCellStyle(sheetName, dealerCell, dealerCell, defaultCellStyle)

		totalRefill := 0
		col := 3
		for _, model := range headers[2 : len(headers)-1] { // Exclude "Total Refill"
			cell, _ := excelize.CoordinatesToCellName(col, row)
			if requiredRefill, exists := modelRefill[model]; exists && requiredRefill > 0 {
				f.SetCellValue(sheetName, cell, requiredRefill)
				f.SetCellStyle(sheetName, cell, cell, refillCellStyle)
				totalRefill += requiredRefill
			} else {
				f.SetCellStyle(sheetName, cell, cell, defaultCellStyle)
			}
			col++
		}

		// Set total refill value
		totalCell, _ := excelize.CoordinatesToCellName(col, row)
		f.SetCellValue(sheetName, totalCell, totalRefill)
		f.SetCellStyle(sheetName, totalCell, totalCell, defaultCellStyle)
		row++
	}

	// Save report to output directory
	outputPath := filepath.Join(outputDir, "ra_norms_report.xlsx")
	excel.AdjustColumnWidths(f, sheetName)
	fmt.Printf("RA report generated successfully: %s\n", outputDir)
	return f.SaveAs(outputPath)

}
