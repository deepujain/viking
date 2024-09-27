package report

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"viking-reports/internal/config"
	"viking-reports/internal/repository"
	"viking-reports/internal/utils"
	"viking-reports/pkg/excel"

	"github.com/xuri/excelize/v2"
)

// ... existing code ...

type SalesTargetGenerator struct {
	cfg *config.Config
	//tseMappingRepo  repository.TSEMappingRepository
	salesTargetRepo repository.SalesTargetRepository
	tseMappingRepo  repository.TSEMappingRepository
}

func NewSalesTargetGenerator(cfg *config.Config) *SalesTargetGenerator {
	tseMappingRepo := repository.NewExcelTSEMappingRepository(cfg.CommonFiles.TSEMapping)
	return &SalesTargetGenerator{
		cfg:             cfg,
		salesTargetRepo: repository.NewExcelSalesTargetRepository(),
		tseMappingRepo:  tseMappingRepo,
	}
}

func (s *SalesTargetGenerator) Generate() error {
	fmt.Println("Generating sales target report for TSE.")

	tseMap, _ := s.tseMappingRepo.GetRetailerCodeToTSEMap()

	fileName := filepath.Base(s.cfg.ReportFiles.SalesReport) // {{ edit_1 }}
	fmt.Printf("** Input: Fetching monthly sales (%s) from Tally **", fileName)

	sales, err := s.salesTargetRepo.ReadSales(s.cfg.ReportFiles.SalesReport, tseMap)
	if err != nil {
		return fmt.Errorf("error : %w", err)
	}
	fmt.Println()
	fmt.Println("\n== Begin processing! ==")
	// Create separate maps for SMART, ACCESSORIES, and others
	// Change maps to slices
	var smartPhoneSales []*repository.SalesData  // {{ edit_1 }}
	var accessoriesSales []*repository.SalesData // {{ edit_1 }}
	var otherSales []*repository.SalesData       // {{ edit_1 }}

	for _, data := range sales {
		if strings.Contains(data.ItemName, "SMART") {
			//fmt.Printf(" %s -> %d -> %s \n", data.ItemName, data.MTDS, data.TSE)
			smartPhoneSales = append(smartPhoneSales, data) // {{ edit_2 }}
		} else if strings.Contains(data.ItemName, "ACCESSORIES") || strings.Contains(data.ItemName, "Buds") {
			accessoriesSales = append(accessoriesSales, data) // {{ edit_2 }}
		} else if strings.Contains(data.ItemName, "Item Name") {
			continue
		} else {
			otherSales = append(otherSales, data) // {{ edit_2 }}
		}
	}

	reportFile := excel.NewFile()
	outputDir := utils.GenerateOutputPath(s.cfg.OutputDir, "sales_report")
	salesTargetSheet := "Sales Target"
	// Create a new sheet
	if _, err := reportFile.NewSheet(salesTargetSheet); err != nil {
		return fmt.Errorf("error creating new sheet: %w", err)
	}

	reportFile.DeleteSheet("Sheet1")
	if err := os.MkdirAll(outputDir, os.ModePerm); err != nil {
		return fmt.Errorf("error creating output directory: %w", err)
	}

	// Invoke writeSalesReport for each category
	fmt.Println("Write monthly sales of SMART PHONES")
	// Create a map for TSE overall targets
	smartPhoneTargets := map[string]int{
		"Krishna": 2490,
		"Sathish": 1900,
		"Harish":  600,
	}
	if err := s.writeSalesTarget(reportFile, salesTargetSheet, smartPhoneSales, smartPhoneTargets, "SMART PHONES", 1); err != nil {
		return fmt.Errorf("error writing smartphone sales report: %w", err)
	}
	fmt.Println()
	fmt.Println("Write monthly sales of ACCESSORIES")
	// Create a map for TSE overall targets
	accessTarget := map[string]int{
		"Krishna": 1000,
		"Sathish": 800,
		"Harish":  600,
	}
	if err := s.writeSalesTarget(reportFile, salesTargetSheet, accessoriesSales, accessTarget, "ACCESSORIES", 8); err != nil {
		return fmt.Errorf("error writing accessories sales report: %w", err)
	}
	fmt.Println()
	fmt.Println("Write monthly sales of OTHERS")
	if err := s.writeSalesTarget(reportFile, salesTargetSheet, otherSales, accessTarget, "OTHERS", 15); err != nil {
		return fmt.Errorf("error writing other sales report: %w", err)
	}
	excel.AdjustColumnWidths(reportFile, salesTargetSheet)
	fileName1 := "sales_report.xlsx"
	outputPath := filepath.Join(outputDir, fileName1)
	reportFile.SaveAs(outputPath)

	fmt.Println("== End processing! ==")
	fmt.Println()
	fmt.Printf("** Output: Sales report generated successfully in: %s ** \n", outputDir)
	return nil
}

func (g *SalesTargetGenerator) writeSalesTarget(f *excelize.File, salesReportSheet string, sales []*repository.SalesData,
	tseSalesTarget map[string]int, productType string, startRow int) error {

	fmt.Printf("Compute and write overall targets for TSE for == %s ==\n", productType)
	if err := excel.WriteHeadersIdx(f, salesReportSheet, []string{productType}, startRow, 5); err != nil {
		return err
	}

	startRow++
	targetHeaders := []string{"TSE", "Target: Overall", "Achieved", "Balance", "Balance %"}
	// Write Overall Target

	overallRow, err := g.writeTarget(sales, tseSalesTarget, f, salesReportSheet, targetHeaders, startRow)
	if err != nil {
		return err
	}

	overallRow++
	return nil
}

func (*SalesTargetGenerator) writeTarget(sales []*repository.SalesData, target map[string]int, f *excelize.File,
	salesReportSheet string, headers []string, startRow int) (int, error) {

	salesAcheivedByTSE := make(map[string]*repository.SalesData)
	for _, data := range sales {
		tse := data.TSE
		if tse != "" {
			if existingData, exists := salesAcheivedByTSE[tse]; exists {
				existingData.MTDS += data.MTDS
				existingData.Value += data.Value
			} else {
				salesAcheivedByTSE[tse] = &repository.SalesData{ // {{ edit_1 }}
					TSE:        tse,
					MTDS:       data.MTDS,
					Value:      data.Value,
					DealerCode: "",
					DealerName: "",
					ItemName:   "",
				}
			}
		}
	}

	if err := excel.WriteHeadersIdx(f, salesReportSheet, headers, startRow, 0); err != nil {
		return 0, err
	}
	targetRow := startRow + 1

	greenStyle, _ := f.NewStyle(&excelize.Style{ // {{ edit_1 }}
		Fill: excelize.Fill{Type: "pattern", Color: []string{"00FF00"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	lightYellowStyle, _ := f.NewStyle(&excelize.Style{ // {{ edit_2 }}
		Fill: excelize.Fill{Type: "pattern", Color: []string{"FFE5B4"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	fmt.Println(headers)
	for _, data := range salesAcheivedByTSE {
		// Print each entry to the console
		fmt.Printf("TSE: %s, MTDS: %d, Value: %d\n", data.TSE, data.MTDS, data.Value) // {{ edit_1 }}
	}

	for _, data := range salesAcheivedByTSE {
		tgt := target[data.TSE]
		bal := target[data.TSE] - data.MTDS
		balPct := (float64(bal) / float64(tgt)) * 100.00

		tseCellData := []interface{}{
			data.TSE,
			tgt,
			data.MTDS,
			bal,
			balPct,
		}
		if err := excel.WriteRow(f, salesReportSheet, targetRow, tseCellData); err != nil {
			return 0, err
		}

		inrFormat := "0.00"
		numberStyle, _ := f.NewStyle(&excelize.Style{
			CustomNumFmt: &inrFormat,
			Border: []excelize.Border{
				{Type: "left", Color: "000000", Style: 1},
				{Type: "top", Color: "000000", Style: 1},
				{Type: "bottom", Color: "000000", Style: 1},
				{Type: "right", Color: "000000", Style: 1},
			},
		})
		col := 4
		cell := fmt.Sprintf("%s%d", string('A'+col), targetRow)
		if err := f.SetCellStyle(salesReportSheet, cell, cell, numberStyle); err != nil {
			return targetRow, fmt.Errorf("error setting style for cell %s: %w", cell, err)
		}

		// Set styles for Achieved and Balance columns
		achievedCell := fmt.Sprintf("%s%d", string('A'+2), targetRow) // Achieved column
		if err := f.SetCellStyle(salesReportSheet, achievedCell, achievedCell, greenStyle); err != nil {
			return targetRow, fmt.Errorf("error setting style for cell %s: %w", achievedCell, err)
		}

		balanceCell := fmt.Sprintf("%s%d", string('A'+3), targetRow) // Balance column
		if err := f.SetCellStyle(salesReportSheet, balanceCell, balanceCell, lightYellowStyle); err != nil {
			return targetRow, fmt.Errorf("error setting style for cell %s: %w", balanceCell, err)
		}
		targetRow++
	}
	return targetRow, nil
}
