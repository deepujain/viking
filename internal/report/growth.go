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

type GrowthReportGenerator struct {
	cfg            *config.Config
	salesRepo      repository.SalesRepository
	tseMappingRepo repository.TSEMappingRepository
}

func NewGrowthReportGenerator(cfg *config.Config) *GrowthReportGenerator {
	return &GrowthReportGenerator{
		cfg:            cfg,
		salesRepo:      repository.NewExcelSalesRepository(), // Use one of the file paths
		tseMappingRepo: repository.NewExcelTSEMappingRepository(cfg.CommonFiles.TSEMapping),
	}
}

func (g *GrowthReportGenerator) Generate() error {
	fmt.Println("Generating Growth report...")

	fmt.Print("Fetching month to date sell out report")
	mtdSOData, err := g.salesRepo.GetSellData(g.cfg.ReportFiles.GrowthReport.MTDSO)
	if err != nil {
		return fmt.Errorf("error reading MTD SO data: %w", err)
	}

	fmt.Print("Fetching last month to date sell out report")
	lmtdSOData, err := g.salesRepo.GetSellData(g.cfg.ReportFiles.GrowthReport.LMTDSO)
	if err != nil {
		return fmt.Errorf("error reading LMTD SO data: %w", err)
	}

	fmt.Print("Fetching month to date sell through report")
	mtdSTData, err := g.salesRepo.GetSellData(g.cfg.ReportFiles.GrowthReport.MTDST)
	if err != nil {
		return fmt.Errorf("error reading MTD ST data: %w", err)
	}

	fmt.Print("Fetching last month to date sell through report")
	lmtdSTData, err := g.salesRepo.GetSellData(g.cfg.ReportFiles.GrowthReport.LMTDST)
	if err != nil {
		return fmt.Errorf("error reading LMTD ST data: %w", err)
	}

	fmt.Println("Fetching retailer code to TSE map.")
	tseMapping, err := g.tseMappingRepo.GetRetailerCodeToTSEMap()
	if err != nil {
		return fmt.Errorf("error reading TSE mapping: %w", err)
	}

	fmt.Println("Start computation of growth report.")
	report := g.generateGrowthReport(mtdSOData, lmtdSOData, mtdSTData, lmtdSTData)
	fmt.Println("Growth report computed.")

	fmt.Println("Write growth report.")
	outputDir := utils.GenerateOutputPath(g.cfg.OutputDir, "growth_report")
	if err := g.writeGrowthReport(outputDir, report, tseMapping); err != nil {
		return fmt.Errorf("error writing growth report: %w", err)
	}

	fmt.Printf("Growth report generated successfully: %s\n", outputDir)
	return nil
}

func (g *GrowthReportGenerator) generateGrowthReport(mtdSOData, lmtdSOData, mtdSTData, lmtdSTData map[string]*repository.SellData) []repository.GrowthData {
	var report []repository.GrowthData

	for dealerCode := range mtdSOData {
		mtdSO := g.getOrCreateSellData(mtdSOData, dealerCode)
		lmtdSO := g.getOrCreateSellData(lmtdSOData, dealerCode)
		mtdST := g.getOrCreateSellData(mtdSTData, dealerCode)
		lmtdST := g.getOrCreateSellData(lmtdSTData, dealerCode)

		reportEntry := repository.GrowthData{
			DealerCode:  dealerCode,
			DealerName:  mtdSO.DealerName,
			MTDSO:       mtdSO.MTDS,
			LMTDSO:      lmtdSO.MTDS,
			GrowthSOPct: utils.CalculateGrowthPercentage(float64(mtdSO.MTDS), float64(lmtdSO.MTDS)), // Convert to float64
			MTDST:       mtdST.MTDS,
			LMTDST:      lmtdST.MTDS,
			GrowthSTPct: utils.CalculateGrowthPercentage(float64(mtdST.MTDS), float64(lmtdST.MTDS)), // Convert to float64
		}

		report = append(report, reportEntry)
	}

	sort.Slice(report, func(i, j int) bool {
		return report[i].GrowthSOPct > report[j].GrowthSOPct
	})
	return report
}

func (g *GrowthReportGenerator) getOrCreateSellData(data map[string]*repository.SellData, dealerCode string) *repository.SellData {
	if data, exists := data[dealerCode]; exists {
		return data
	}
	return &repository.SellData{DealerCode: dealerCode, DealerName: "", MTDS: 0}
}

func (g *GrowthReportGenerator) writeGrowthReport(outputDir string, report []repository.GrowthData, tseMapping map[string]string) error {
	f := excel.NewFile()
	sheetName := "Growth Report"

	// Create a new sheet
	if _, err := f.NewSheet(sheetName); err != nil {
		return fmt.Errorf("error creating new sheet: %w", err)
	}
	f.DeleteSheet("Sheet1")

	if err := os.MkdirAll(outputDir, os.ModePerm); err != nil {
		return fmt.Errorf("error creating output directory: %w", err)
	}

	headers := []string{"TSE", "Dealer Code", "Dealer Name", "MTD SO", "LMTD SO", "Growth SO %", "MTD ST", "LMTD ST", "Growth ST %"}
	if err := excel.WriteHeaders(f, sheetName, headers); err != nil {
		return err
	}

	// Sort report by TSE using tseMapping
	sort.Slice(report, func(i, j int) bool {
		return tseMapping[report[i].DealerCode] < tseMapping[report[j].DealerCode]
	})

	row := 2
	for _, entry := range report {
		cellData := []interface{}{
			tseMapping[entry.DealerCode],
			entry.DealerCode,
			entry.DealerName,
			entry.MTDSO,
			entry.LMTDSO,
			entry.GrowthSOPct,
			entry.MTDST,
			entry.LMTDST,
			entry.GrowthSTPct,
		}
		if err := excel.WriteRow(f, sheetName, row, cellData); err != nil {
			return err
		}

		//Apply all kinds of styles
		// Create a new style that inherits from numberStyle and adds background fill
		redStyle, _ := f.NewStyle(&excelize.Style{
			Fill: excelize.Fill{Type: "pattern", Color: []string{"FF0000"}, Pattern: 1}, // Light red background
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
		// Create a new style that inherits from numberStyle and adds background fill
		orangeStyle, _ := f.NewStyle(&excelize.Style{
			Fill: excelize.Fill{Type: "pattern", Color: []string{"FFFF00"}, Pattern: 1}, // Light red background
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
		// Create a new style that inherits from numberStyle and adds background fill
		greenStyle, _ := f.NewStyle(&excelize.Style{
			Fill: excelize.Fill{Type: "pattern", Color: []string{"00FF00"}, Pattern: 1}, // Light red background
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

		// Apply background color based on GrowthSOPct
		if entry.GrowthSOPct < -60 {
			f.SetCellStyle(sheetName, fmt.Sprintf("F%d", row), fmt.Sprintf("F%d", row), redStyle) // Assuming redStyle is defined
		} else if entry.GrowthSOPct >= -59 && entry.GrowthSOPct < 0 {
			f.SetCellStyle(sheetName, fmt.Sprintf("F%d", row), fmt.Sprintf("F%d", row), orangeStyle) // Assuming orangeStyle is defined
		} else if entry.GrowthSOPct > 0 {
			f.SetCellStyle(sheetName, fmt.Sprintf("F%d", row), fmt.Sprintf("F%d", row), greenStyle) // Assuming greenStyle is defined
		}

		// Apply background color based on GrowthSTPct
		if entry.GrowthSTPct < -60 {
			f.SetCellStyle(sheetName, fmt.Sprintf("I%d", row), fmt.Sprintf("I%d", row), redStyle) // Assuming redStyle is defined
		} else if entry.GrowthSTPct >= -59 && entry.GrowthSTPct < 0 {
			f.SetCellStyle(sheetName, fmt.Sprintf("I%d", row), fmt.Sprintf("I%d", row), orangeStyle) // Assuming orangeStyle is defined
		} else if entry.GrowthSTPct > 0 {
			f.SetCellStyle(sheetName, fmt.Sprintf("I%d", row), fmt.Sprintf("I%d", row), greenStyle) // Assuming greenStyle is defined
		}

		row++
	}

	// Ensure the output path has a valid extension
	fileName := "sales_growth_report.xlsx"
	outputPath := filepath.Join(outputDir, fileName)
	excel.AdjustColumnWidths(f, sheetName)
	return f.SaveAs(outputPath)
}
