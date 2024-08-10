package utils

import (
	"fmt"
	"sort"

	"github.com/xuri/excelize/v2"
)

// Write categorized data to an Excel file
func WriteCategorizedData(dirPath string, fileName string, categorizedData map[string]map[string]interface{}, retailerNameToCodeMap map[string]string) error {
	// Convert the map to a slice for sorting
	var dataSlice []PartyData
	for retailerName, amounts := range categorizedData {
		dataSlice = append(dataSlice, PartyData{
			RetailerName: retailerName,
			Amounts:      amounts,
		})
	}

	// Sort the slice by the "Total" column in descending order
	sort.Slice(dataSlice, func(i, j int) bool {
		return dataSlice[i].Amounts["Total"].(float64) > dataSlice[j].Amounts["Total"].(float64)
	})

	// Create a new Excel file for the output
	f := excelize.NewFile()
	sheetName := "Sheet1"
	if _, err := f.NewSheet(sheetName); err != nil {
		return fmt.Errorf("failed to create new sheet: %w", err)
	}

	// Define styles
	headerStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"FFFF00"}, // Yellow background
			Pattern: 1,
		},
		Font: &excelize.Font{
			Bold: true,
		},
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

	cellStyle, err := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create cell style: %w", err)
	}

	// Write the header row
	headers := []string{"Retailer Code", "Retailer Name", "0-7 days", "8-14 days", "15-21 days", "22-30 days", ">30 days", "Total", "TSE"}
	for i, header := range headers {
		cell := fmt.Sprintf("%s%d", string('A'+i), 1)
		if err := f.SetCellValue(sheetName, cell, header); err != nil {
			return fmt.Errorf("failed to set header cell value: %w", err)
		}
		if err := f.SetCellStyle(sheetName, cell, cell, headerStyle); err != nil {
			return fmt.Errorf("failed to set header cell style: %w", err)
		}
	}

	// Adjust column widths
	f.SetColWidth(sheetName, "A", "A", 15) // Width for Retailer Code
	f.SetColWidth(sheetName, "B", "B", 46) // Width for Retailer Name
	for i := 2; i < len(headers); i++ {
		col := string('A' + i)
		f.SetColWidth(sheetName, col, col, 13) // Width for other columns
	}

	// Write the sorted categorized data to the output file
	for rowNum, data := range dataSlice {
		rowIndex := rowNum + 2
		retailerCode, exists := retailerNameToCodeMap[data.RetailerName]
		if !exists {
			retailerCode = "N/A" // Default value if retailer code is not found
		}
		if err := f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), retailerCode); err != nil {
			return fmt.Errorf("failed to set retailer code: %w", err)
		}
		if err := f.SetCellStyle(sheetName, fmt.Sprintf("A%d", rowIndex), fmt.Sprintf("A%d", rowIndex), cellStyle); err != nil {
			return fmt.Errorf("failed to set cell style: %w", err)
		}
		if err := f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIndex), data.RetailerName); err != nil {
			return fmt.Errorf("failed to set retailer name: %w", err)
		}
		if err := f.SetCellStyle(sheetName, fmt.Sprintf("B%d", rowIndex), fmt.Sprintf("B%d", rowIndex), cellStyle); err != nil {
			return fmt.Errorf("failed to set cell style: %w", err)
		}
		for i, category := range headers[2 : len(headers)-1] {
			cell := fmt.Sprintf("%s%d", string('C'+i), rowIndex)
			if err := f.SetCellValue(sheetName, cell, data.Amounts[category].(float64)); err != nil {
				return fmt.Errorf("failed to set category value: %w", err)
			}
			if err := f.SetCellStyle(sheetName, cell, cell, numberStyle); err != nil {
				return fmt.Errorf("failed to set number style: %w", err)
			}
		}
		if err := f.SetCellValue(sheetName, fmt.Sprintf("I%d", rowIndex), data.Amounts["TSE"].(string)); err != nil {
			return fmt.Errorf("failed to set TSE value: %w", err)
		}
		if err := f.SetCellStyle(sheetName, fmt.Sprintf("I%d", rowIndex), fmt.Sprintf("I%d", rowIndex), cellStyle); err != nil {
			return fmt.Errorf("failed to set cell style: %w", err)
		}
	}

	// Save the output file
	filePath := fmt.Sprintf("%s/%s", dirPath, fileName)
	if err := f.SaveAs(filePath); err != nil {
		return fmt.Errorf("failed to save output file: %w", err)
	}

	return nil
}
