package report

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"
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
	fmt.Printf("Generating Sales Target report for %s %d \n", time.Now().Month().String(), time.Now().Year())

	fmt.Println("** Input: Fetching retailer code to TSE name map from metadata. **")
	tseMap, _ := s.tseMappingRepo.GetRetailerCodeToTSEMap()

	fmt.Print("** Input: Fetching monthly sales from Tally and computing sales for each retailer **")
	sales, err := s.salesTargetRepo.ComputeSales(s.cfg.ReportFiles.SalesReport, tseMap)
	if err != nil {
		return fmt.Errorf("error : %w", err)
	}

	fmt.Println("Generating output...")
	// Create separate maps for SMART, ACCESSORIES, and others
	smartPhoneSales := make(map[string]*repository.SalesData)
	accessoriesSales := make(map[string]*repository.SalesData)
	otherSales := make(map[string]*repository.SalesData)

	for key, data := range sales {

		if strings.Contains(data.ItemName, "SMART") {
			smartPhoneSales[key] = data
		} else if strings.Contains(data.ItemName, "ACCESSORIES") || strings.Contains(data.ItemName, "Buds") {
			accessoriesSales[key] = data
		} else if strings.Contains(data.ItemName, "Item Name") {
			continue
		} else {
			otherSales[key] = data
		}
	}

	reportFile := excel.NewFile()
	outputDir := utils.GenerateOutputPath(s.cfg.OutputDir, "sales_report")
	// Invoke writeSalesReport for each category
	fmt.Println("Write monthly sales of SMART PHONES for each retailer")
	// Create a map for TSE overall targets
	tseSmartPhoneSalesOverallTargets := map[string]int{
		"Krishna Murthy": 2490,
		"SATHISH":        1900,
		"HARISH":         600,
	}
	if err := s.writeSalesReport(reportFile, outputDir, smartPhoneSales, tseSmartPhoneSalesOverallTargets, "SMART PHONES"); err != nil {
		return fmt.Errorf("error writing smartphone sales report: %w", err)
	}

	fmt.Println("Write monthly sales of ACCESSORIES for each retailer")
	// Create a map for TSE overall targets
	tseAccessOverallTargets := map[string]int{
		"Krishna Murthy": 1000,
		"SATHISH":        800,
		"HARISH":         600,
	}
	if err := s.writeSalesReport(reportFile, outputDir, accessoriesSales, tseAccessOverallTargets, "ACCESSORIES"); err != nil {
		return fmt.Errorf("error writing accessories sales report: %w", err)
	}

	fmt.Println("Write monthly sales of OTHERS for each retailer")
	if err := s.writeSalesReport(reportFile, outputDir, otherSales, tseSmartPhoneSalesOverallTargets, "OTHERS"); err != nil {
		return fmt.Errorf("error writing other sales report: %w", err)
	}
	fmt.Printf("Sales report generated successfully for %s %d: %s \n", time.Now().Month().String(), time.Now().Year(), outputDir)
	return nil
}

func (g *SalesTargetGenerator) writeSalesReport(f *excelize.File, outputDir string, sales map[string]*repository.SalesData, tseSalesTarget map[string]int, productType string) error {
	salesReportSheet := productType
	// Create a new sheet
	if _, err := f.NewSheet(salesReportSheet); err != nil {
		return fmt.Errorf("error creating new sheet: %w", err)
	}

	f.DeleteSheet("Sheet1")
	if err := os.MkdirAll(outputDir, os.ModePerm); err != nil {
		return fmt.Errorf("error creating output directory: %w", err)
	}

	fmt.Printf("Compute and write overall targets for TSE for %s \n", productType)
	var startRow = 1
	if err := excel.WriteHeadersIdx(f, salesReportSheet, []string{productType}, startRow, 5); err != nil {
		return err
	}
	startRow++
	if err := excel.WriteHeadersIdx(f, salesReportSheet, []string{"TSE Targets"}, startRow, 5); err != nil {
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
	if err := excel.WriteHeadersIdx(f, salesReportSheet, []string{"Sales"}, overallRow, 5); err != nil {
		return err
	}
	fmt.Printf("Compute and write sales report of each retailer and TSE for %s.\n", productType)
	err = g.writeSales(overallRow+1, f, salesReportSheet, sales)
	if err != nil {
		return err
	}

	excel.AdjustColumnWidths(f, salesReportSheet)
	fileName := "sales_report.xlsx"
	outputPath := filepath.Join(outputDir, fileName)
	return f.SaveAs(outputPath)
}

func (*SalesTargetGenerator) writeTarget(sales map[string]*repository.SalesData, target map[string]int, f *excelize.File, salesReportSheet string, headers []string, startRow int) (int, error) {
	salesTSE := make(map[string]*repository.SalesData)
	for _, data := range sales {
		tse := data.TSE
		if tse != "" {
			if existingData, exists := salesTSE[tse]; exists {
				existingData.MTDS += data.MTDS
				existingData.Value += data.Value
			} else {
				salesTSE[tse] = &repository.SalesData{ // {{ edit_1 }}
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

	for _, data := range salesTSE {
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

func (*SalesTargetGenerator) writeSales(row int, f *excelize.File, salesReportSheet string, sales map[string]*repository.SalesData) error {
	headers := []string{"Dealer Code", "Dealer Name", "Sell Out", "Total Sales Value(₹)", "TSE"}
	if err := excel.WriteHeadersIdx(f, salesReportSheet, headers, row, 0); err != nil {
		return err
	}

	inrFormat := "#,##,##0.00"
	numberStyle, _ := f.NewStyle(&excelize.Style{
		CustomNumFmt: &inrFormat,
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})

	redStyle, _ := f.NewStyle(&excelize.Style{
		Fill:         excelize.Fill{Type: "pattern", Color: []string{"FF9999"}, Pattern: 1},
		CustomNumFmt: &inrFormat,
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
		Font: &excelize.Font{
			Bold: true,
		},
	})

	salesSlice := make([]*repository.SalesData, 0, len(sales))
	for _, data := range sales {
		salesSlice = append(salesSlice, data)
	}

	sort.Slice(salesSlice, func(i, j int) bool {
		if salesSlice[i].TSE == salesSlice[j].TSE {
			return salesSlice[i].Value > salesSlice[j].Value
		}
		return salesSlice[i].TSE > salesSlice[j].TSE
	})

	totalQty := 0
	totalValue := 0
	salesRow := row + 1
	for _, data := range salesSlice {
		cellData := []interface{}{
			data.DealerCode,
			data.DealerName,
			data.MTDS,
			data.Value,
			data.TSE,
		}
		totalQty += data.MTDS
		totalValue += data.Value

		if err := excel.WriteRow(f, salesReportSheet, salesRow, cellData); err != nil {
			return err
		}

		for col := 3; col <= 3; col++ {
			cell := fmt.Sprintf("%s%d", string('A'+col), salesRow)
			var style int
			if data.MTDS < 0 {
				style = redStyle
			} else {
				style = numberStyle
			}
			if err := f.SetCellStyle(salesReportSheet, cell, cell, style); err != nil {
				return fmt.Errorf("error setting style for cell %s: %w", cell, err)
			}
		}
		salesRow++
	}

	totalRow := salesRow
	totalCellData := []interface{}{
		"Total",
		"",
		totalQty,
		totalValue,
		"",
	}
	if err := excel.WriteRow(f, salesReportSheet, totalRow, totalCellData); err != nil {
		return err
	}

	for col := 3; col <= 3; col++ {
		cell := fmt.Sprintf("%s%d", string('A'+col), totalRow)
		if err := f.SetCellStyle(salesReportSheet, cell, cell, numberStyle); err != nil {
			return fmt.Errorf("error setting style for cell %s: %w", cell, err)
		}
	}
	return nil
}
