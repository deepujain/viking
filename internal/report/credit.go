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

type CreditReportGenerator struct {
	cfg            *config.Config
	creditRepo     repository.CreditRepository
	debitRepo      repository.DebitRepository
	inventoryRepo  repository.InventoryRepository
	tseMappingRepo repository.TSEMappingRepository
}

func NewCreditReportGenerator(cfg *config.Config) *CreditReportGenerator {
	priceRepo := repository.NewExcelProductPriceRepository(cfg.CommonFiles.PriceList)
	tseMappingRepo := repository.NewExcelTSEMappingRepository(cfg.CommonFiles.TSEMapping)

	priceData, _ := priceRepo.GetProductPrices()
	tseMapping, _ := tseMappingRepo.GetRetailerCodeToTSEMap()

	return &CreditReportGenerator{
		cfg:            cfg,
		creditRepo:     repository.NewExcelCreditRepository(cfg.ReportFiles.CreditReport.Bills),
		debitRepo:      repository.NewExcelDebitRepository(cfg.ReportFiles.DebitReport.Debits),
		inventoryRepo:  repository.NewExcelInventoryRepository(cfg.ReportFiles.InventoryReport, priceData, tseMapping),
		tseMappingRepo: repository.NewExcelTSEMappingRepository(cfg.CommonFiles.TSEMapping),
	}
}

func (g *CreditReportGenerator) Generate() error {

	bills, err := g.creditRepo.GetBills()
	if err != nil {
		return fmt.Errorf("error reading bills: %w", err)
	}

	tseMapping, err := g.tseMappingRepo.GetRetailerNameToTSEMap("Tally Name(Dealer Name)")
	if err != nil {
		return fmt.Errorf("error reading TSE mapping: %w", err)
	}

	retailerNameToCodeMap, err := g.tseMappingRepo.GetRetailerNameToCodeMap()
	if err != nil {
		return fmt.Errorf("error reading Name to Code mapping: %w", err)
	}

	retailerNameToDebitMap, err := g.debitRepo.GetDebit()
	if err != nil {
		return fmt.Errorf("error reading Name to debit mapping: %w", err)
	}
	fmt.Println("** Input: Fetching the stock inventory of retailers from DMS portal  **")

	fmt.Println("\n== Begin processing! ==")
	retailerCredit := g.creditRepo.AggregateCreditByRetailer(bills, tseMapping, retailerNameToCodeMap)

	inventoryData, err := g.inventoryRepo.ComputeInventoryShortFall()
	if err != nil { // Check for error
		return fmt.Errorf("error computing inventory shortfall: %w", err) // Handle the error
	}

	outputDir := utils.GenerateOutputPath(g.cfg.OutputDir, "credit_reports")
	if err := g.writeCreditReports(outputDir, retailerCredit, inventoryData, retailerNameToDebitMap); err != nil {
		return fmt.Errorf("error writing credit reports: %w", err)
	}
	fmt.Println("== End processing! ==")
	fmt.Println()
	fmt.Printf("** Output: Credit reports generated successfully in: %s ** \n", outputDir)
	return nil
}

func (g *CreditReportGenerator) writeCreditReports(outputDir string, retailerCredit map[string]map[string]interface{},
	inventoryData map[string]*repository.InventoryShortFallRepo, retailerNameToDebitMap map[string]float64) error {
	totalDealerCreditWithTSE := make(map[string]map[string]map[string]interface{})
	totalDealerCreditMissingTSE := make(map[string]map[string]interface{})

	for retailerName, credit := range retailerCredit {
		tseName := credit["TSE"].(string)
		if tseName == "" {
			totalDealerCreditMissingTSE[retailerName] = credit
		} else {
			if totalDealerCreditWithTSE[tseName] == nil {
				totalDealerCreditWithTSE[tseName] = make(map[string]map[string]interface{})
			}
			totalDealerCreditWithTSE[tseName][retailerName] = credit
		}
	}
	for tseName, retailerCredit := range totalDealerCreditWithTSE {
		fmt.Printf("Generating total credit report for %d retailers assigned to %s \n", len(retailerCredit), tseName)
		fileName := fmt.Sprintf("%s_credit_report.xlsx", tseName)
		if err := g.writeCreditReport(outputDir, fileName, retailerCredit, inventoryData, retailerNameToDebitMap); err != nil {
			return fmt.Errorf("error writing file for TSE %s: %w", tseName, err)
		}
	}

	if len(totalDealerCreditMissingTSE) > 0 {
		fmt.Printf("Generating total credit report for %d retailers for which TSE's are *not* assigned!  \n", len(totalDealerCreditMissingTSE))
		if err := g.writeCreditReport(outputDir, "TSE_MISSING_credit_report.xlsx", totalDealerCreditMissingTSE, inventoryData, retailerNameToDebitMap); err != nil {
			return fmt.Errorf("error writing TSE_MISSING file: %w", err)
		}
	}

	return nil
}

func (g *CreditReportGenerator) writeCreditReport(outputDir, fileName string, data map[string]map[string]interface{},
	inventoryData map[string]*repository.InventoryShortFallRepo, retailerNameToDebitMap map[string]float64) error {
	f := excel.NewFile()
	sheetName := "Credit Report"
	// Create a new sheet
	if _, err := f.NewSheet(sheetName); err != nil {
		return fmt.Errorf("error creating new sheet: %w", err)
	}
	f.DeleteSheet("Sheet1")

	if err := os.MkdirAll(outputDir, os.ModePerm); err != nil {
		return fmt.Errorf("error creating output directory: %w", err)
	}

	// Custom number format for Indian numbering
	inrFormat := "#,##,##0"
	numberStyle, err := f.NewStyle(&excelize.Style{
		CustomNumFmt: &inrFormat, // Custom number format for Indian numbering
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create number style: %w", err)
	}

	headers := []string{"Retailer Code", "Retailer Name", "Received: Last 1 days (₹)", "Credit: 0-7 Days(₹)", "Credit: 8-14 Days(₹)", "Credit: 15-20 Days(₹)", "Credit: 21-30 Days(₹)", "Credit: 31+ Days(₹)", "Total Credit(₹)", "Total Inventory Cost(₹)", "Inventory Shortfall (₹)", "TSE"}
	if err := excel.WriteHeaders(f, sheetName, headers); err != nil {
		return err
	}

	row := 2
	inventoryShortfalls := make([]struct {
		Credit    map[string]interface{}
		Shortfall float64
	}, 0)

	for _, retailerCredit := range data {
		inventoryCost := 0.0
		if dealerData, exists := inventoryData[retailerCredit["Retailer Code"].(string)]; exists { // Fetch inventory cost using retailer code
			inventoryCost = dealerData.TotalInventoryCost
		} else {
			fmt.Printf("Inventory Cost missing for dealer '%s' !\n", retailerCredit["Retailer Code"])
		}
		inventoryShortFall := inventoryCost - retailerCredit["Total Credit"].(float64)

		// Store the retailer credit and its shortfall
		inventoryShortfalls = append(inventoryShortfalls, struct {
			Credit    map[string]interface{}
			Shortfall float64
		}{Credit: retailerCredit, Shortfall: inventoryShortFall})
	}

	// Sort by inventoryShortFall (ascending)
	sort.Slice(inventoryShortfalls, func(i, j int) bool {
		return inventoryShortfalls[i].Shortfall < inventoryShortfalls[j].Shortfall
	})

	// Write sorted data to the sheet
	var totalDebitLastDay = 0.0
	for _, item := range inventoryShortfalls {
		retailerCredit := item.Credit
		inventoryShortFall := item.Shortfall
		inventoryCost := 0.0
		if dealerData, exists := inventoryData[retailerCredit["Retailer Code"].(string)]; exists {
			inventoryCost = dealerData.TotalInventoryCost
		}
		totalDebitLastDay = totalDebitLastDay + retailerNameToDebitMap[retailerCredit["Retailer Name"].(string)]
		cellData := []interface{}{
			retailerCredit["Retailer Code"],
			retailerCredit["Retailer Name"],
			retailerNameToDebitMap[retailerCredit["Retailer Name"].(string)],
			retailerCredit["0-7 Days"],
			retailerCredit["8-14 Days"],
			retailerCredit["15-20 Days"],
			retailerCredit["21-30 Days"],
			retailerCredit["31+ Days"],
			retailerCredit["Total Credit"],
			inventoryCost, // Ensure inventoryCost is defined before use
			inventoryShortFall,
			retailerCredit["TSE"],
		}

		if err := excel.WriteRow(f, sheetName, row, cellData); err != nil {
			return err
		}

		//Apply all kinds of styles

		// Ensure numeric values are of type float64 to all numeric columns
		for i := 2; i <= 9; i++ { // Columns C (2) to I (7)
			if val, ok := retailerCredit[headers[i]].(float64); ok {
				cellData[i-2] = val // Replace with float64 value
			}
		}

		// Apply number style to all numeric columns
		for col := 2; col <= 9; col++ { // Columns C (3) to I (8)
			cell := fmt.Sprintf("%s%d", string('A'+col), row) // Convert column index to letter
			if err := f.SetCellStyle(sheetName, cell, cell, numberStyle); err != nil {
				return fmt.Errorf("error setting style for cell %s: %w", cell, err)
			}
		}

		// Create a red Style
		redStyle, err := f.NewStyle(&excelize.Style{
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
		if err != nil {
			return fmt.Errorf("failed to create background style: %w", err)
		}

		// Apply backgroundStyle style to 22-30 Days(₹)	and 31+ Days(₹)
		for col := 6; col <= 7; col++ {
			cell := fmt.Sprintf("%s%d", string('A'+col), row) // Convert column index to letter
			if err := f.SetCellStyle(sheetName, cell, cell, redStyle); err != nil {
				return fmt.Errorf("error setting style for cell %s: %w", cell, err)
			}
		}

		// Apply redStyle to Inventory Shortfall cells if negative
		if inventoryShortFall < 0 {
			cell := fmt.Sprintf("K%d", row) // Column J for Inventory Shortfall
			if err := f.SetCellStyle(sheetName, cell, cell, redStyle); err != nil {
				return fmt.Errorf("error setting style for cell %s: %w", cell, err)
			}
		}
		row++
	}

	// Calculate totals
	total := make([]float64, 8) // For columns 0-7 Days to Total Credit
	for _, item := range inventoryShortfalls {
		retailerCredit := item.Credit
		total[0] += retailerCredit["0-7 Days"].(float64)
		total[1] += retailerCredit["8-14 Days"].(float64)
		total[2] += retailerCredit["15-20 Days"].(float64)
		total[3] += retailerCredit["21-30 Days"].(float64)
		total[4] += retailerCredit["31+ Days"].(float64)
		total[5] += retailerCredit["Total Credit"].(float64)

		inventoryCost := 0.0
		if dealerData, exists := inventoryData[retailerCredit["Retailer Code"].(string)]; exists {
			inventoryCost = dealerData.TotalInventoryCost
		}
		total[6] += inventoryCost                                            // Total Inventory Cost
		total[7] += retailerCredit["Total Credit"].(float64) - inventoryCost // Total Inventory Shortfall
	}

	// Write totals to the last row
	totalCellData := []interface{}{
		"Total", // Label for the total row
		"",      // Retailer Name
		totalDebitLastDay,
		total[0], total[1], total[2], total[3], total[4], total[5], total[6], total[7], "",
	}
	if err := excel.WriteRow(f, sheetName, row, totalCellData); err != nil {
		return err
	}

	// Apply number style to total row
	for col := 2; col <= 10; col++ { // Columns C (4) to I (9)
		cell := fmt.Sprintf("%s%d", string('A'+col), row) // Convert column index to letter
		if err := f.SetCellStyle(sheetName, cell, cell, numberStyle); err != nil {
			return fmt.Errorf("error setting style for cell %s: %w", cell, err)
		}
	}

	excel.AdjustColumnWidths(f, sheetName)
	outputPath := filepath.Join(outputDir, fileName)
	return f.SaveAs(outputPath)
}
