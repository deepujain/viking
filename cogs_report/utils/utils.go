// Package utils provides utility functions for generating inventory cost reports.
package utils

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

// InventoryData represents the inventory data for each dealer.
type InventoryData struct {
	DealerCode         string
	DealerName         string
	TotalInventoryCost float64
	TSE                string
}

// getColumnIndex returns the index of a column given its name.
func getColumnIndex(f *excelize.File, sheetName, columnName string) (int, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return -1, fmt.Errorf("failed to read rows: %w", err)
	}
	if len(rows) == 0 {
		return -1, fmt.Errorf("no rows found in sheet")
	}
	for i, col := range rows[0] {
		if col == columnName {
			return i, nil
		}
	}
	return -1, fmt.Errorf("column %s not found", columnName)
}

// readPriceData reads the price data (NLC, MOP) from an Excel file.
func readPriceData(filePath string) (map[string]float64, error) {
	xlFile, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file %s: %w", filePath, err)
	}
	defer xlFile.Close()

	sheetName := xlFile.GetSheetName(0)
	rows, err := xlFile.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows from sheet: %w", err)
	}

	materialCodeIdx, err := getColumnIndex(xlFile, sheetName, "Material Code")
	if err != nil {
		return nil, fmt.Errorf("failed to get Material Code column index: %w", err)
	}
	nlcIdx, err := getColumnIndex(xlFile, sheetName, "NLC")
	if err != nil {
		return nil, fmt.Errorf("failed to get NLC column index: %w", err)
	}

	priceData := make(map[string]float64)
	for _, row := range rows[1:] { // Skip header row
		if materialCode := row[materialCodeIdx]; materialCode != "" {
			priceData[materialCode] = parseFloat(row[nlcIdx])
		}
	}

	return priceData, nil
}

// readInventoryData reads the inventory data from an Excel file and calculates total inventory cost.
// readInventoryData reads the inventory data from an Excel file and calculates total inventory cost.
func readInventoryData(filePath string, priceData map[string]float64, tseMapping map[string]string) ([]InventoryData, error) {
	xlFile, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file %s: %w", filePath, err)
	}
	defer xlFile.Close()

	sheetName := xlFile.GetSheetName(0)
	rows, err := xlFile.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows from sheet: %w", err)
	}

	materialCodeIdx, err := getColumnIndex(xlFile, sheetName, "Material Code")
	if err != nil {
		return nil, fmt.Errorf("failed to get Material Code column index: %w", err)
	}
	dealerCodeIdx, err := getColumnIndex(xlFile, sheetName, "Dealer Code")
	if err != nil {
		return nil, fmt.Errorf("failed to get Dealer Code column index: %w", err)
	}
	dealerNameIdx, err := getColumnIndex(xlFile, sheetName, "Dealer Name")
	if err != nil {
		return nil, fmt.Errorf("failed to get Dealer Name column index: %w", err)
	}

	inventoryData := make(map[string]*InventoryData)
	for _, row := range rows[1:] { // Skip header row
		materialCode := row[materialCodeIdx]
		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]
		tse := tseMapping[dealerCode]

		if materialCode == "" || dealerCode == "" {
			continue
		}

		price := priceData[materialCode]
		if data, exists := inventoryData[dealerCode]; exists {
			data.TotalInventoryCost += price
		} else {
			inventoryData[dealerCode] = &InventoryData{
				DealerCode:         dealerCode,
				DealerName:         dealerName,
				TotalInventoryCost: price,
				TSE:                tse,
			}
		}
	}

	// Convert map to slice
	inventoryDataSlice := make([]InventoryData, 0, len(inventoryData))
	for _, data := range inventoryData {
		inventoryDataSlice = append(inventoryDataSlice, *data)
	}

	// Sort slice by TotalInventoryCost in descending order
	sort.Slice(inventoryDataSlice, func(i, j int) bool {
		return inventoryDataSlice[i].TotalInventoryCost > inventoryDataSlice[j].TotalInventoryCost
	})

	return inventoryDataSlice, nil
}

// parseFloat converts a string to a float64, returning 0 if the conversion fails.
func parseFloat(str string) float64 {
	value, err := strconv.ParseFloat(str, 64)
	if err != nil {
		return 0
	}
	return value
}

// adjustColumnWidths adjusts the widths of columns in the Excel sheet.
func adjustColumnWidths(f *excelize.File, sheetName string) {
	columnWidths := map[string]float64{
		"A": 13, // Dealer Code
		"B": 46, // Dealer Name
		"C": 22, // Total Inventory Cost
		"D": 15, //TSE
	}

	for col, width := range columnWidths {
		f.SetColWidth(sheetName, col, col, width)
	}
}

// writeHeaders writes the headers to the Excel sheet.
func writeHeaders(f *excelize.File, sheetName string) error {
	headerStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"FFFF00"}, Pattern: 1},
		Font: &excelize.Font{Bold: true},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create header style: %w", err)
	}

	headers := []string{"Dealer Code", "Dealer Name", "Total Inventory Cost (₹)", "TSE"}
	for i, header := range headers {
		cell := fmt.Sprintf("%s%d", string('A'+i), 1)
		if err := f.SetCellValue(sheetName, cell, header); err != nil {
			return fmt.Errorf("failed to set header cell value: %w", err)
		}
		if err := f.SetCellStyle(sheetName, cell, cell, headerStyle); err != nil {
			return fmt.Errorf("failed to set header cell style: %w", err)
		}
	}

	return nil
}

// writeSingleRow writes a single row of data to the Excel sheet.
func writeSingleRow(f *excelize.File, sheetName string, rowIndex int, data InventoryData) error {
	// Custom number format for Indian numbering
	inrFormat := "#,##,##0.00"
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

	// Set cell values
	cellStyle, err := f.NewStyle(&excelize.Style{
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
	dealerCodeCell := fmt.Sprintf("A%d", rowIndex)
	dealerNameCell := fmt.Sprintf("B%d", rowIndex)
	f.SetCellValue(sheetName, dealerCodeCell, data.DealerCode)
	f.SetCellStyle(sheetName, dealerCodeCell, dealerCodeCell, cellStyle)
	f.SetCellValue(sheetName, dealerNameCell, data.DealerName)
	f.SetCellStyle(sheetName, dealerNameCell, dealerNameCell, cellStyle)
	tSECELL := fmt.Sprintf("D%d", rowIndex)
	f.SetCellValue(sheetName, tSECELL, data.TSE)
	f.SetCellStyle(sheetName, tSECELL, tSECELL, cellStyle)
	// Set inventory cost with number style
	costCell := fmt.Sprintf("C%d", rowIndex)
	f.SetCellValue(sheetName, costCell, data.TotalInventoryCost)
	f.SetCellStyle(sheetName, costCell, costCell, numberStyle)

	return nil
}

// writeInventoryReport writes the inventory report to an Excel file.
func writeInventoryReport(f *excelize.File, sheetName string, inventoryData []InventoryData) error {
	// Create a new sheet
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to create new sheet: %w", err)
	}
	f.SetActiveSheet(index)

	if err := writeHeaders(f, sheetName); err != nil {
		return err
	}

	// Write each row of inventory data
	for i, data := range inventoryData {
		if err := writeSingleRow(f, sheetName, i+2, data); err != nil {
			return err
		}
	}

	adjustColumnWidths(f, sheetName)

	return nil
}

// readTSEToRetailerMapping reads the TSE to retailer mapping from an Excel file.
func readTSEToRetailerMapping(tseMappingFilePath string) (map[string]string, error) {
	tseMappingFile, err := excelize.OpenFile(tseMappingFilePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open TSE mapping file: %w", err)
	}
	defer tseMappingFile.Close()

	tseSheetNames := tseMappingFile.GetSheetList()
	if len(tseSheetNames) == 0 {
		return nil, fmt.Errorf("no sheets found in the TSE mapping file")
	}

	tseSheet := tseSheetNames[0]
	tseRows, err := tseMappingFile.GetRows(tseSheet)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows from TSE mapping file: %w", err)
	}

	tseMapping := make(map[string]string)
	for _, row := range tseRows[1:] { // Skip header row
		if len(row) < 16 {
			continue
		}
		dealerCode, tseName := row[5], row[15]
		if dealerCode != "" && tseName != "" {
			tseMapping[dealerCode] = tseName
		}
	}
	return tseMapping, nil
}

// RunInventoryCostReport runs the entire inventory cost report generation process.
func RunInventoryCostReport() error {
	dataDir := "../data"
	priceFile := filepath.Join(dataDir, "ProductPriceList.xlsx")
	inventoryFile := filepath.Join(dataDir, "DealerInventory.xlsx")
	tseMappingFile := filepath.Join(dataDir, "VIKING'S - DEALER Credit Period LIST.xlsx")

	today := time.Now().Format("2006-01-02")
	dirPath := fmt.Sprintf("./daily_growth_reports_%s", today)
	if err := os.MkdirAll(dirPath, os.ModePerm); err != nil {
		return fmt.Errorf("failed to create directory: %w", err)
	}

	outputPath := filepath.Join(dirPath, "daily_inventory_cost_report.xlsx")

	priceData, err := readPriceData(priceFile)
	if err != nil {
		return fmt.Errorf("error reading price data: %w", err)
	}

	tseMapping, err := readTSEToRetailerMapping(tseMappingFile)
	if err != nil {
		return fmt.Errorf("error reading TSE mapping: %w", err)
	}

	inventoryData, err := readInventoryData(inventoryFile, priceData, tseMapping)
	if err != nil {
		return fmt.Errorf("error reading inventory data: %w", err)
	}

	// Create a new Excel file
	f := excelize.NewFile()
	sheetName := "Inventory Report"

	if err := writeInventoryReport(f, sheetName, inventoryData); err != nil {
		return fmt.Errorf("error writing inventory report: %w", err)
	}

	if err := f.SaveAs(outputPath); err != nil {
		return fmt.Errorf("failed to save file: %w", err)
	}

	fmt.Printf("Inventory cost report generated successfully: %s\n", outputPath)
	return nil
}
