package utils

import (
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

// Bill struct holds the details of each bill
type Bill struct {
	Date          string
	RefNo         string
	RetailerName  string
	PendingAmount float64
	DueDate       string
	AgeOfBill     int
}

// PartyData struct to hold the categorized data along with the party name
type PartyData struct {
	RetailerName string
	Amounts      map[string]interface{}
}

// Open the input Excel file
// Print sheet names to verify
// Use the first sheet or specify the sheet name if known
// Assumes the first sheet is the one you want
// Read rows from the sheet
// Process the data rows, ignoring the last row which contains totals
// Start from row 12 (index 11)
// Skip last row (total row)
func readBills(inputFilePath string) []Bill {
	xlFile, err := excelize.OpenFile(inputFilePath)
	if err != nil {
		log.Fatalf("Failed to open input file: %v", err)
	}

	sheetNames := xlFile.GetSheetList()
	if len(sheetNames) == 0 {
		log.Fatalf("No sheets found in the input file")
	}

	sheet := sheetNames[0]

	rows, err := xlFile.GetRows(sheet)
	if err != nil {
		log.Fatalf("Failed to get rows: %v", err)
	}

	var bills []Bill
	totalRows := len(rows)
	for i, row := range rows[11:] {
		if len(row) < 6 || i+11 == totalRows-1 {
			continue
		}

		date := row[0]
		refNo := row[1]
		retailerName := row[2]
		pendingAmount, err := strconv.ParseFloat(row[3], 64)
		if err != nil {
			log.Printf("Error parsing pending amount: %v", err)
			continue
		}
		dueDate := row[4]
		ageOfBill, err := strconv.Atoi(row[5])
		if err != nil {
			log.Printf("Error parsing age of bill: %v", err)
			continue
		}

		bills = append(bills, Bill{
			Date:          date,
			RefNo:         refNo,
			RetailerName:  retailerName,
			PendingAmount: pendingAmount,
			DueDate:       dueDate,
			AgeOfBill:     ageOfBill,
		})
	}
	return bills
}

// Read TSE mapping file
// Print sheet names to verify
// Assumes the first sheet is the one you want
// Read rows from the TSE sheet
// Assuming first row is headers
// Dealer Name in column 7 (index 6)
// TSE Name in column 16 (index 15)

func readTSEToRetailerMapping(tseMappingFilePath string) map[string]string {
	tseMappingFile, err := excelize.OpenFile(tseMappingFilePath)
	if err != nil {
		log.Fatalf("Failed to open TSE mapping file: %v", err)
	}

	tseSheetNames := tseMappingFile.GetSheetList()
	if len(tseSheetNames) == 0 {
		log.Fatalf("No sheets found in the TSE mapping file")
	}

	tseSheet := tseSheetNames[0]

	tseRows, err := tseMappingFile.GetRows(tseSheet)
	if err != nil {
		log.Fatalf("Failed to get rows from TSE mapping file: %v", err)
	}

	tseMapping := make(map[string]string)
	for _, row := range tseRows[1:] {
		if len(row) < 16 {
			continue
		}

		dealerName := row[6]
		tseName := row[15]

		if dealerName != "" && tseName != "" {
			tseMapping[dealerName] = tseName
		}
	}
	return tseMapping
}

// aggregateCreditByRetailer categorizes the bills by retailer and age category
func aggregateCreditByRetailer(bills []Bill, tseMapping map[string]string) map[string]map[string]interface{} {
	ageCategories := []string{"0-7 days", "8-14 days", "15-21 days", "22-30 days", ">30 days"}
	aggregatedCreditRetailer := make(map[string]map[string]interface{})

	// First, group all bills by retailer name
	groupedBills := make(map[string][]Bill)
	for _, bill := range bills {
		groupedBills[bill.RetailerName] = append(groupedBills[bill.RetailerName], bill)
	}

	// Now process each retailer's bills
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
			switch {
			case bill.AgeOfBill >= 0 && bill.AgeOfBill <= 7:
				aggregatedCreditRetailer[retailerName]["0-7 days"] = aggregatedCreditRetailer[retailerName]["0-7 days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill >= 8 && bill.AgeOfBill <= 14:
				aggregatedCreditRetailer[retailerName]["8-14 days"] = aggregatedCreditRetailer[retailerName]["8-14 days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill >= 15 && bill.AgeOfBill <= 21:
				aggregatedCreditRetailer[retailerName]["15-21 days"] = aggregatedCreditRetailer[retailerName]["15-21 days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill >= 22 && bill.AgeOfBill <= 30:
				aggregatedCreditRetailer[retailerName]["22-30 days"] = aggregatedCreditRetailer[retailerName]["22-30 days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill > 30:
				aggregatedCreditRetailer[retailerName][">30 days"] = aggregatedCreditRetailer[retailerName][">30 days"].(float64) + bill.PendingAmount
			}
		}
		aggregatedCreditRetailer[retailerName]["Total"] = totalPendingAmount
		aggregatedCreditRetailer[retailerName]["TSE"] = tseMapping[retailerName] // Add TSE name
	}

	return aggregatedCreditRetailer
}

func writeCategorizedData(dirPath string, fileName string, categorizedData map[string]map[string]interface{}) error {
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
	headers := []string{"Retailer Name", "0-7 days", "8-14 days", "15-21 days", "22-30 days", ">30 days", "Total", "TSE"}
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
	f.SetColWidth(sheetName, "A", "A", 46) // Width for Retailer Name
	for i := 1; i < len(headers); i++ {
		col := string('A' + i)
		f.SetColWidth(sheetName, col, col, 13) // Width for other columns
	}

	// Write the sorted categorized data to the output file
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

	// Save the output file
	filePath := fmt.Sprintf("%s/%s", dirPath, fileName)
	if err := f.SaveAs(filePath); err != nil {
		return fmt.Errorf("failed to save output file: %w", err)
	}

	return nil
}

func RunCreditReport() {
	inputFilePath := "../data/Bills.xlsx"                                     // Read from current directory
	tseMappingFilePath := "../data/VIKING'S - DEALER Credit Period LIST.xlsx" // Read the TSE mapping file

	// Get today's date for folder name
	today := time.Now().Format("2006-01-02")
	dirPath := fmt.Sprintf("./daily_credit_reports_%s", today)

	// Create directory for today's date
	if err := os.MkdirAll(dirPath, os.ModePerm); err != nil {
		log.Fatalf("Failed to create directory: %v", err)
	}

	bills := readBills(inputFilePath)

	tseMapping := readTSEToRetailerMapping(tseMappingFilePath)

	// Categorize bills by retailer and age category
	aggregatedData := aggregateCreditByRetailer(bills, tseMapping)

	// Separate data by TSE
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
}
