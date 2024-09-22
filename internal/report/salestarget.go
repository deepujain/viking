package report

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
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
	fmt.Println("Generating Sales Target report.")

	fmt.Print("Fetching monthly sales from Tally and computing sales for each retailer")
	sales, err := s.salesTargetRepo.ComputeSales(s.cfg.ReportFiles.SalesReport)
	if err != nil {
		return fmt.Errorf("error : %w", err)
	}

	tseMap, _ := s.tseMappingRepo.GetRetailerCodeToTSEMap()
	fmt.Println(tseMap)
	// Print each key-value pair in the sales map
	for key, value := range sales {
		fmt.Printf("Key: %s, Value: %+v\n", key, value)
	}

	// Use priceData, tseMapping, and creditData in your COGS calculation logic here
	reportFile := excel.NewFile()

	outputDir := utils.GenerateOutputPath(s.cfg.OutputDir, "sales_report")
	if err := s.writeSalesReport(reportFile, outputDir, sales, tseMap); err != nil {
		return fmt.Errorf("error writing sales report: %w", err)
	}
	return nil
}

func (g *SalesTargetGenerator) writeSalesReport(f *excelize.File, outputDir string, sales map[string]*repository.SalesData, tseMap map[string]string) error {
	salesReportSheet := fmt.Sprintf("Sales: %s %d", time.Now().Month().String(), time.Now().Year()) // Updated to include month and year
	// Create a new sheet
	if _, err := f.NewSheet(salesReportSheet); err != nil {
		return fmt.Errorf("error creating new sheet: %w", err)
	}

	f.DeleteSheet("Sheet1")
	if err := os.MkdirAll(outputDir, os.ModePerm); err != nil {
		return fmt.Errorf("error creating output directory: %w", err)
	}

	headers := []string{"Dealer Code", "Dealer Name", "QTY", "Total Sales Value(₹)", "TSE"}
	if err := excel.WriteHeaders(f, salesReportSheet, headers); err != nil {
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
	inventorySlice := make([]*repository.SalesData, 0, len(sales))
	for _, data := range sales {
		inventorySlice = append(inventorySlice, data)
	}

	// Sort by TSE and then by Cost-Credit Difference
	sort.Slice(inventorySlice, func(i, j int) bool {
		if inventorySlice[i].TSE == inventorySlice[j].TSE {
			return inventorySlice[i].Value > inventorySlice[j].Value
		}
		return inventorySlice[i].TSE < inventorySlice[j].TSE
	})

	row := 2
	for _, data := range inventorySlice {
		cellData := []interface{}{
			data.DealerCode,
			data.DealerName,
			data.MTDS,
			data.Value,
			tseMap[data.DealerCode],
		}
		if err := excel.WriteRow(f, salesReportSheet, row, cellData); err != nil {
			return err
		}

		// Apply number style to numeric columns (0-7 Days to Total Credit)
		for col := 2; col <= 3; col++ { // Columns C (3) to I (5)
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
	excel.AdjustColumnWidths(f, salesReportSheet)
	fileName := "sales_report.xlsx"
	outputPath := filepath.Join(outputDir, fileName)
	return f.SaveAs(outputPath)
}
