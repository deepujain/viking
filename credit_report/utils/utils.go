// Package utils provides utility functions for generating credit reports.
package utils

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

// Bill represents the details of each bill.
type Bill struct {
	Date          string
	RefNo         string
	RetailerName  string
	PendingAmount float64
	DueDate       string
	AgeOfBill     int
}

// PartyData represents the categorized data for a retailer.
type PartyData struct {
	RetailerName string
	Amounts      map[string]interface{}
}

// readBills reads and parses bill data from an Excel file.
func readBills(inputFilePath string) ([]Bill, error) {
	xlFile, err := excelize.OpenFile(inputFilePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open input file: %w", err)
	}
	defer xlFile.Close()

	sheetNames := xlFile.GetSheetList()
	if len(sheetNames) == 0 {
		return nil, fmt.Errorf("no sheets found in the input file")
	}

	sheet := sheetNames[0]
	rows, err := xlFile.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %w", err)
	}

	var bills []Bill
	totalRows := len(rows)
	for i, row := range rows[11:] { // Start from row 12 (index 11)
		if len(row) < 6 || i+11 == totalRows-1 { // Skip last row (total row)
			continue
		}

		pendingAmount, err := strconv.ParseFloat(row[3], 64)
		if err != nil {
			log.Printf("Error parsing pending amount: %v", err)
			continue
		}
		ageOfBill, err := strconv.Atoi(row[5])
		if err != nil {
			log.Printf("Error parsing age of bill: %v", err)
			continue
		}

		bills = append(bills, Bill{
			Date:          row[0],
			RefNo:         row[1],
			RetailerName:  row[2],
			PendingAmount: pendingAmount,
			DueDate:       row[4],
			AgeOfBill:     ageOfBill,
		})
	}
	return bills, nil
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
		dealerName, tseName := row[6], row[15]
		if dealerName != "" && tseName != "" {
			tseMapping[dealerName] = tseName
		}
	}
	return tseMapping, nil
}

// aggregateCreditByRetailer categorizes the bills by retailer and age category.
func aggregateCreditByRetailer(bills []Bill, tseMapping map[string]string) map[string]map[string]interface{} {
	ageCategories := []string{"0-7 days", "8-14 days", "15-21 days", "22-30 days", ">30 days"}
	aggregatedCreditRetailer := make(map[string]map[string]interface{})

	// Group bills by retailer name
	groupedBills := make(map[string][]Bill)
	for _, bill := range bills {
		groupedBills[bill.RetailerName] = append(groupedBills[bill.RetailerName], bill)
	}

	// Process each retailer's bills
	for retailerName, partyBills := range groupedBills {
		aggregatedCreditRetailer[retailerName] = make(map[string]interface{})
		totalPendingAmount := 0.0

		// Initialize categories with 0.0
		for _, category := range ageCategories {
			aggregatedCreditRetailer[retailerName][category] = 0.0
		}

		// Categorize bills into age categories and compute total pending amount
		for _, bill := range partyBills {
			totalPendingAmount += bill.PendingAmount
			category := getAgeCategory(bill.AgeOfBill)
			aggregatedCreditRetailer[retailerName][category] = aggregatedCreditRetailer[retailerName][category].(float64) + bill.PendingAmount
		}
		aggregatedCreditRetailer[retailerName]["Total"] = totalPendingAmount
		aggregatedCreditRetailer[retailerName]["TSE"] = tseMapping[retailerName] // Add TSE name
	}

	return aggregatedCreditRetailer
}

// getAgeCategory returns the age category for a given bill age.
func getAgeCategory(age int) string {
	switch {
	case age >= 0 && age <= 7:
		return "0-7 days"
	case age >= 8 && age <= 14:
		return "8-14 days"
	case age >= 15 && age <= 21:
		return "15-21 days"
	case age >= 22 && age <= 30:
		return "22-30 days"
	default:
		return ">30 days"
	}
}

// writeCategorizedData writes the categorized data to an Excel file.
func writeCategorizedData(dirPath, fileName string, categorizedData map[string]map[string]interface{}) error {
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
	headerStyle, numberStyle, cellStyle, err := createStyles(f)
	if err != nil {
		return err
	}

	// Write the header row
	headers := []string{"Retailer Name", "0-7 days", "8-14 days", "15-21 days", "22-30 days", ">30 days", "Total", "TSE"}
	if err := writeHeaders(f, sheetName, headers, headerStyle); err != nil {
		return err
	}

	// Adjust column widths
	adjustColumnWidths(f, sheetName, headers)

	// Write the sorted categorized data to the output file
	if err := writeData(f, sheetName, dataSlice, headers, numberStyle, cellStyle); err != nil {
		return err
	}

	// Save the output file
	filePath := filepath.Join(dirPath, fileName)
	if err := f.SaveAs(filePath); err != nil {
		return fmt.Errorf("failed to save output file: %w", err)
	}

	return nil
}

// createStyles creates the styles used in the Excel file.
func createStyles(f *excelize.File) (int, int, int, error) {
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
		return 0, 0, 0, fmt.Errorf("failed to create header style: %w", err)
	}

	inrFormat := "#,##,##0.00"
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
		return 0, 0, 0, fmt.Errorf("failed to create number style: %w", err)
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
		return 0, 0, 0, fmt.Errorf("failed to create cell style: %w", err)
	}

	return headerStyle, numberStyle, cellStyle, nil
}

// writeHeaders writes the header row to the Excel file.
func writeHeaders(f *excelize.File, sheetName string, headers []string, headerStyle int) error {
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

// adjustColumnWidths adjusts the column widths in the Excel file.
func adjustColumnWidths(f *excelize.File, sheetName string, headers []string) {
	f.SetColWidth(sheetName, "A", "A", 46) // Width for Retailer Name
	for i := 1; i < len(headers); i++ {
		col := string('A' + i)
		f.SetColWidth(sheetName, col, col, 13) // Width for other columns
	}
}

// writeData writes the categorized data to the Excel file.
func writeData(f *excelize.File, sheetName string, dataSlice []PartyData, headers []string, numberStyle, cellStyle int) error {
	for rowNum, data := range dataSlice {
		rowIndex := rowNum + 2
		if err := f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), data.RetailerName); err != nil {
			return fmt.Errorf("failed to set retailer name: %w", err)
		}
		if err := f.SetCellStyle(sheetName, fmt.Sprintf("A%d", rowIndex), fmt.Sprintf("A%d", rowIndex), cellStyle); err != nil {
			return fmt.Errorf("failed to set cell style: %w", err)
		}
		for i, category := range headers[1 : len(headers)-1] {
			cell := fmt.Sprintf("%s%d", string('B'+i), rowIndex)
			if err := f.SetCellValue(sheetName, cell, data.Amounts[category].(float64)); err != nil {
				return fmt.Errorf("failed to set category value: %w", err)
			}
			if err := f.SetCellStyle(sheetName, cell, cell, numberStyle); err != nil {
				return fmt.Errorf("failed to set number style: %w", err)
			}
		}
		if err := f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowIndex), data.Amounts["TSE"].(string)); err != nil {
			return fmt.Errorf("failed to set TSE value: %w", err)
		}
		if err := f.SetCellStyle(sheetName, fmt.Sprintf("H%d", rowIndex), fmt.Sprintf("H%d", rowIndex), cellStyle); err != nil {
			return fmt.Errorf("failed to set cell style: %w", err)
		}
	}
	return nil
}

// RunCreditReport runs the entire credit report generation process.
func RunCreditReport() error {
	inputFilePath := "../data/Bills.xlsx"
	tseMappingFilePath := "../data/VIKING'S - DEALER Credit Period LIST.xlsx"

	// Get today's date for folder name
	today := time.Now().Format("2006-01-02")
	dirPath := fmt.Sprintf("./daily_credit_reports_%s", today)

	// Create directory for today's date
	if err := os.MkdirAll(dirPath, os.ModePerm); err != nil {
		return fmt.Errorf("failed to create directory: %w", err)
	}

	// Read bills from input file
	bills, err := readBills(inputFilePath)
	if err != nil {
		return fmt.Errorf("failed to read bills: %w", err)
	}

	// Read TSE mapping
	tseMapping, err := readTSEToRetailerMapping(tseMappingFilePath)
	if err != nil {
		return fmt.Errorf("failed to read TSE mapping: %w", err)
	}

	// Categorize bills by retailer and age category
	aggregatedData := aggregateCreditByRetailer(bills, tseMapping)

	// Separate data by TSE
	tseFiles, missingTSEData := separateDataByTSE(aggregatedData)

	// Write output files for each TSE
	for tseName, data := range tseFiles {
		fileName := fmt.Sprintf("%s_credit_report.xlsx", tseName)
		if err := writeCategorizedData(dirPath, fileName, data); err != nil {
			log.Printf("Error writing file for TSE %s: %v", tseName, err)
		} else {
			fmt.Printf("Credit report for TSE %s saved to: %s/%s\n", tseName, dirPath, fileName)
		}
	}

	// Write output file for missing TSE
	if len(missingTSEData) > 0 {
		if err := writeCategorizedData(dirPath, "TSE_MISSING_credit_report.xlsx", missingTSEData); err != nil {
			log.Printf("Error writing TSE_MISSING file: %v", err)
		} else {
			fmt.Printf("Credit report for missing TSE saved to: %s/TSE_MISSING_credit_report.xlsx\n", dirPath)
		}
	}

	return nil
}

// separateDataByTSE separates the aggregated data by TSE.
func separateDataByTSE(aggregatedData map[string]map[string]interface{}) (map[string]map[string]map[string]interface{}, map[string]map[string]interface{}) {
	tseFiles := make(map[string]map[string]map[string]interface{})
	missingTSEData := make(map[string]map[string]interface{})

	for retailerName, amounts := range aggregatedData {
		tseName := amounts["TSE"].(string)
		if tseName == "" {
			missingTSEData[retailerName] = amounts
		} else {
			if tseFiles[tseName] == nil {
				tseFiles[tseName] = make(map[string]map[string]interface{})
			}
			tseFiles[tseName][retailerName] = amounts
		}
	}

	return tseFiles, missingTSEData
}
