package report

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"
	"viking-reports/internal/config"
	"viking-reports/internal/repository"
	"viking-reports/internal/utils"
	"viking-reports/pkg/excel"

	"github.com/xuri/excelize/v2"
)

type PriceListGenerator struct {
	cfg           *config.Config
	priceListRepo repository.PriceListRepository
}

func NewPriceListGenerator(cfg *config.Config) *PriceListGenerator {
	return &PriceListGenerator{
		cfg:           cfg,
		priceListRepo: repository.NewExcelPriceListRepository(cfg.ReportFiles.PriceListFile, cfg.ReportFiles.InventoryReport),
	}
}

func (p *PriceListGenerator) Generate() error {
	currentMonthYear := time.Now().Format("January 2006")
	fmt.Printf("\nGenerating flat price list of SKUs for the month of %s \n", currentMonthYear)
	priceData, err := p.priceListRepo.GetPriceListData()
	if err != nil {
		return err
	}
	fmt.Printf("Price list generated successfully, total SKUs: %d\n", len(priceData))
	fmt.Printf("Read inventory data to generate material code for the month of %s \n", currentMonthYear)

	materialCodeMap, err := p.priceListRepo.GetMaterialCodeMap()
	if err != nil {
		return err
	}
	fmt.Printf("Material code map generated successfully, Size: %d\n", len(materialCodeMap))

	outputDir := utils.GenerateMonthlyOutputPath(p.cfg.OutputDir, "price_list")
	p.writePriceList(outputDir, priceData, materialCodeMap)
	fmt.Printf("\nPrice list written successfully in: %s\n", outputDir)
	return nil
}

func (p *PriceListGenerator) writePriceList(outputDir string, priceData []repository.PriceListRow, materialCodeMap map[string]int) error {
	f := excel.NewFile()
	sheetName := "Price List"

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
		CustomNumFmt: &inrFormat,
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

	// Write headers
	headers := []string{"Type", "Model", "Color", "Variant", "NLC", "MOP", "MRP", "Material Code"} // Adjust headers as needed
	if err := excel.WriteHeaders(f, sheetName, headers); err != nil {
		return err
	}

	// Write data rows
	for rowIndex, item := range priceData {
		// Create a unique key based on item.Model, item.Color, and combined Storage and Memory
		key := fmt.Sprintf("%s|%s|%s", item.Model, item.Color, (item.Storage + " " + item.Memory))
		cellData := []interface{}{
			item.Type,
			item.Model,
			item.Color,
			item.Storage + " " + item.Memory,
			item.NLC,
			item.Mop,
			item.Mrp,
			fmt.Sprintf("%d", materialCodeMap[strings.ToLower(key)]),
		}

		if err := excel.WriteRow(f, sheetName, rowIndex+2, cellData); err != nil {
			return err
		}

		// Apply number style to price column
		cell := fmt.Sprintf("C%d", rowIndex+2) // Assuming Price is in column C
		if err := f.SetCellStyle(sheetName, cell, cell, numberStyle); err != nil {
			return fmt.Errorf("error setting style for cell %s: %w", cell, err)
		}
	}

	excel.AdjustColumnWidths(f, sheetName)
	outputPath := filepath.Join(outputDir, "price_list.xlsx")
	return f.SaveAs(outputPath)
}
