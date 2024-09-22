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

	fmt.Println("Fetching retailer code to TSE name map from metadata.")
	tseMap, _ := s.tseMappingRepo.GetRetailerCodeToTSEMap()

	fmt.Print("Fetching monthly sales from Tally and computing sales for each retailer")
	sales, err := s.salesTargetRepo.ComputeSales(s.cfg.ReportFiles.SalesReport, tseMap)
	if err != nil {
		return fmt.Errorf("error : %w", err)
	}

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
	if err := s.writeSalesReport(reportFile, outputDir, smartPhoneSales, "SMART PHONES"); err != nil {
		return fmt.Errorf("error writing smartphone sales report: %w", err)
	}
	fmt.Println("Write monthly sales of ACCESSORIES for each retailer")
	if err := s.writeSalesReport(reportFile, outputDir, accessoriesSales, "ACCESSORIES"); err != nil {
		return fmt.Errorf("error writing accessories sales report: %w", err)
	}
	fmt.Println("Write monthly sales of OTHERS for each retailer")
	if err := s.writeSalesReport(reportFile, outputDir, otherSales, "OTHERS"); err != nil {
		return fmt.Errorf("error writing other sales report: %w", err)
	}
	fmt.Printf("Sales report generated successfully for %s %d: %s \n", time.Now().Month().String(), time.Now().Year(), outputDir)
	return nil
}

func (g *SalesTargetGenerator) writeSalesReport(f *excelize.File, outputDir string, sales map[string]*repository.SalesData, productType string) error {
	salesReportSheet := productType
	// Create a new sheet
	if _, err := f.NewSheet(salesReportSheet); err != nil {
		return fmt.Errorf("error creating new sheet: %w", err)
	}

	f.DeleteSheet("Sheet1")
	if err := os.MkdirAll(outputDir, os.ModePerm); err != nil {
		return fmt.Errorf("error creating output directory: %w", err)
	}

	headers := []string{"Dealer Code", "Dealer Name", "Sell Out", "Total Sales Value(₹)", "TSE"}
	if err := excel.WriteHeaders(f, salesReportSheet, headers); err != nil {
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
	salesSlice := make([]*repository.SalesData, 0, len(sales))
	for _, data := range sales {
		salesSlice = append(salesSlice, data)
	}

	// Sort by TSE and then by Total Sales Value(₹)
	sort.Slice(salesSlice, func(i, j int) bool {
		if salesSlice[i].TSE == salesSlice[j].TSE {
			return salesSlice[i].Value > salesSlice[j].Value
		}
		return salesSlice[i].TSE > salesSlice[j].TSE
	})
	// Calculate totals
	totalQty := 0
	totalValue := 0
	row := 2
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

		if err := excel.WriteRow(f, salesReportSheet, row, cellData); err != nil {
			return err
		}

		// Apply number style to numeric columns
		for col := 3; col <= 3; col++ {
			cell := fmt.Sprintf("%s%d", string('A'+col), row) // Convert column index to letter
			var style int
			if data.MTDS < 0 { // Check if MTDS is negative
				style = redStyle // Use redStyle for negative values
			} else {
				style = numberStyle // Use numberStyle for non-negative values
			}
			if err := f.SetCellStyle(salesReportSheet, cell, cell, style); err != nil {
				return fmt.Errorf("error setting style for cell %s: %w", cell, err)
			}
		}
		row++
	}

	// Write totals to the last row
	totalRow := row
	totalCellData := []interface{}{
		"Total", // Label for the total row
		"",      // Dealer Name
		totalQty,
		totalValue,
		"", // TSE
	}
	if err := excel.WriteRow(f, salesReportSheet, totalRow, totalCellData); err != nil {
		return err
	}

	// Apply number style to total row
	for col := 2; col <= 3; col++ { // Columns C (3) to D (4)
		cell := fmt.Sprintf("%s%d", string('A'+col), totalRow) // Convert column index to letter
		if err := f.SetCellStyle(salesReportSheet, cell, cell, numberStyle); err != nil {
			return fmt.Errorf("error setting style for cell %s: %w", cell, err)
		}
	}
	excel.AdjustColumnWidths(f, salesReportSheet)
	fileName := "sales_report.xlsx"
	outputPath := filepath.Join(outputDir, fileName)
	return f.SaveAs(outputPath)
}
