package utils

import (
	"fmt"
	"log"
	"os"
	"sort"
	"time"

	"github.com/xuri/excelize/v2"
)

// SellOutData struct to hold the sell-out data for each retailer
type SellData struct {
	DealerCode string
	DealerName string
	MTDS       int
	LMTDS      int
}

// GrowthReport struct to hold the final report data
type GrowthReport struct {
	DealerCode  string
	DealerName  string
	MTDSO       int
	LMTDSO      int
	GrowthSOPct float64
	MTDST       int
	LMTDST      int
	GrowthSTPct float64
}

// Function to get column index from a given column name
func getColumnIndex(f *excelize.File, sheetName, columnName string) (int, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return -1, fmt.Errorf("failed to read rows: %v", err)
	}
	if len(rows) == 0 {
		return -1, fmt.Errorf("no rows found in sheet")
	}
	headerRow := rows[0]
	for i, col := range headerRow {
		if col == columnName {
			return i, nil
		}
	}
	return -1, fmt.Errorf("column %s not found", columnName)
}

// Function to read Dealer Information and return a map of Dealer Code to Dealer Name
func readDealerInformation(filePath string) (map[string]string, error) {
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

	// Get the column indices for Dealer Code and Dealer Name
	dealerCodeIdx, err := getColumnIndex(xlFile, sheetName, "Dealer Code")
	if err != nil {
		return nil, fmt.Errorf("failed to get Dealer Code column index: %w", err)
	}
	dealerNameIdx, err := getColumnIndex(xlFile, sheetName, "Dealer Name")
	if err != nil {
		return nil, fmt.Errorf("failed to get Dealer Name column index: %w", err)
	}

	// Create a map to store Dealer Code to Dealer Name mapping
	dealerInfo := make(map[string]string)

	for i, row := range rows {
		if i == 0 { // Skip header row
			continue
		}

		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]

		// Skip rows with empty or nil Dealer Code
		if dealerCode == "" {
			log.Printf("Skipping row %d due to empty or nil Dealer Code", i+1)
			continue
		}

		// Store Dealer Code to Dealer Name mapping
		dealerInfo[dealerCode] = dealerName
	}

	return dealerInfo, nil
}

// Function to read Sell-Out and Sell-Through Data from the Excel file by column names
func readSellData(filePath string, dataType string) (map[string]*SellData, error) {
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

	// Map the required column names to their indices based on dataType
	var activateTimeIdx, dealerCodeIdx, dealerNameIdx int
	var errActivateTime, errDealerCode, errDealerName error

	switch dataType {
	case "SO":
		activateTimeIdx, errActivateTime = getColumnIndex(xlFile, sheetName, "Activate Time")
		dealerCodeIdx, errDealerCode = getColumnIndex(xlFile, sheetName, "Dealer Code")
		dealerNameIdx, errDealerName = getColumnIndex(xlFile, sheetName, "Dealer Name")
	case "ST":
		activateTimeIdx, errActivateTime = getColumnIndex(xlFile, sheetName, "transferTime")
		dealerCodeIdx, errDealerCode = getColumnIndex(xlFile, sheetName, "toDealerCode")
		dealerNameIdx, errDealerName = getColumnIndex(xlFile, sheetName, "toDealerName")
	default:
		return nil, fmt.Errorf("unknown dataType: %s", dataType)
	}

	if errActivateTime != nil || errDealerCode != nil || errDealerName != nil {
		return nil, fmt.Errorf("failed to get column indices: %v, %v, %v", errActivateTime, errDealerCode, errDealerName)
	}

	sellData := make(map[string]*SellData)
	currentDate := time.Now() // Get today's date
	currentDay := currentDate.Day()

	for i, row := range rows {
		if i == 0 { // Skip header row
			continue
		}

		activateTimeStr := row[activateTimeIdx]
		activateTime, err := time.Parse("2006-01-02 15:04:05", activateTimeStr)
		if err != nil {
			continue
		}
		// Convert both dates to the same format to compare only the day portion
		activateDay := activateTime.Day()

		// Filter out rows where the day of the month is greater than today's day of the month
		if activateDay > currentDay {
			continue
		}

		dealerCode := row[dealerCodeIdx]
		// Filter out rows where dealerCode is nil or blank
		if dealerCode == "" {
			continue
		}

		dealerName := row[dealerNameIdx]

		if data, exists := sellData[dealerCode]; exists {
			data.MTDS++
		} else {
			sellData[dealerCode] = &SellData{
				DealerCode: dealerCode,
				DealerName: dealerName,
				MTDS:       1,
			}
		}

	}

	return sellData, nil
}

// Function to generate the growth report
func generateGrowthReport(mtdSOData, lmtdSOData, mtdSTData, lmtdSTData map[string]*SellData, dealerInfo map[string]string) []GrowthReport {

	var report []GrowthReport

	// Always prioritize reading from the mtdSTData (Sell-Through) dataset, rather than mtdSOData (Sell-Out).
	// Dealers may have taken SKUs from the distributor (Sell-Through) but might not have sold them to customers yet (Sell-Out).
	// Therefore, it's crucial to focus on Sell-Through data for accurate reporting.
	for dealerCode, dealerName := range dealerInfo {
		lmtdSO := lmtdSOData[dealerCode]
		mtdSO := mtdSOData[dealerCode]
		mtdST := mtdSTData[dealerCode]
		lmtdST := lmtdSTData[dealerCode]

		// If any of the SellData objects is nil, initialize them with zero
		if mtdSO == nil {
			mtdSO = &SellData{DealerCode: dealerCode, DealerName: dealerName, MTDS: 0, LMTDS: 0}
		}
		if lmtdSO == nil {
			lmtdSO = &SellData{DealerCode: dealerCode, DealerName: dealerName, MTDS: 0, LMTDS: 0}
		}
		if mtdST == nil {
			mtdST = &SellData{DealerCode: dealerCode, DealerName: dealerName, MTDS: 0, LMTDS: 0}
		}
		if lmtdST == nil {
			lmtdST = &SellData{DealerCode: dealerCode, DealerName: dealerName, MTDS: 0, LMTDS: 0}
		}

		// Compute growth percentage for Sell-Out
		growthSOPct := 0.0
		if lmtdSO.MTDS > 0 {
			growthSOPct = float64(mtdSO.MTDS-lmtdSO.MTDS) / float64(lmtdSO.MTDS) * 100
		} else if mtdSO.MTDS > 0 {
			// If lmtdSO.MTDS is zero and mtdSO.MTDS is greater than zero
			growthSOPct = 100.0 // or another large number to indicate exceptional growth
		} else {
			// Both lmtdSO.MTDS and mtdSO.MTDS are zero
			growthSOPct = 0.0
		}

		// Compute growth percentage for Sell-Through
		growthSTPct := 0.0
		if lmtdST.MTDS > 0 {
			growthSTPct = float64(mtdST.MTDS-lmtdST.MTDS) / float64(lmtdST.MTDS) * 100
		} else if mtdST.MTDS > 0 {
			// If lmtdST.MTDS is zero and mtdST.MTDS is greater than zero
			growthSTPct = 100.0 // or another large number to indicate exceptional growth
		} else {
			// Both lmtdST.MTDS and mtdST.MTDS are zero
			growthSTPct = 0.0
		}

		reportEntry := GrowthReport{
			DealerCode:  dealerCode,
			DealerName:  mtdSO.DealerName,
			MTDSO:       mtdSO.MTDS,
			LMTDSO:      lmtdSO.MTDS,
			GrowthSOPct: growthSOPct,
			MTDST:       mtdST.MTDS,
			LMTDST:      lmtdST.MTDS,
			GrowthSTPct: growthSTPct,
		}
		//log.Printf("Adding report entry: %+v", reportEntry)
		report = append(report, reportEntry)
	}

	// Sort the report by GrowthSOPct in descending order
	sort.Slice(report, func(i, j int) bool {
		return report[j].GrowthSOPct > report[i].GrowthSOPct
	})

	return report
}

// Function to write the growth report to an Excel file
// Function to write the growth report to an Excel file
func writeGrowthReport(report []GrowthReport, tseMapping map[string]string, outputPath string) error {
	f := excelize.NewFile()
	sheetName := "Growth Report"
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to create new sheet: %w", err)
	}
	f.SetActiveSheet(index)

	// Define styles
	headerStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{"FFFF00"}, // Yellow background
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
	// Write headers
	headers := []string{"Retailer Code", "Retailer Name", "MTD SO", "LMTD SO", "Growth SO (%)", "MTD ST", "LMTD ST", "Growth ST (%)", "TSE"}
	for i, header := range headers {
		cell := fmt.Sprintf("%s%d", string('A'+i), 1)
		if err := f.SetCellValue(sheetName, cell, header); err != nil {
			return fmt.Errorf("failed to set header cell value: %w", err)
		}
		if err := f.SetCellStyle(sheetName, cell, cell, headerStyle); err != nil {
			return fmt.Errorf("failed to set header cell style: %w", err)
		}
	}

	// Style for Growth SO (%) based on value
	positiveGrowthStyle, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Color: "000000", // Black text color
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"00FF00"}, // Green background for positive growth
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create positive growth style: %w", err)
	}

	negativeGrowthStyle, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Color: "000000", // Black text color
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"FF0000"}, // Green background for positive growth
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create negative growth style: %w", err)
	}

	// Write data
	for row, data := range report {
		rowIndex := row + 2
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), data.DealerCode)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIndex), data.DealerName)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowIndex), data.MTDSO)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowIndex), data.LMTDSO)
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowIndex), data.GrowthSOPct)
		f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowIndex), data.MTDST)
		f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowIndex), data.LMTDST)
		f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowIndex), data.GrowthSTPct)
		f.SetCellValue(sheetName, fmt.Sprintf("I%d", rowIndex), tseMapping[data.DealerCode])

		// Format Growth SO (%) as whole number
		growthSOPct := int(data.GrowthSOPct) // Convert to whole number
		gsoCell := fmt.Sprintf("E%d", rowIndex)
		f.SetCellValue(sheetName, gsoCell, growthSOPct)

		// Apply style for Growth SO (%) based on value
		gsoStyle := positiveGrowthStyle
		if data.GrowthSOPct < 0 {
			gsoStyle = negativeGrowthStyle
		}
		f.SetCellStyle(sheetName, gsoCell, gsoCell, gsoStyle)

		// Format Growth ST (%) as whole number
		growthSTPct := int(data.GrowthSTPct) // Convert to whole number
		gstCell := fmt.Sprintf("H%d", rowIndex)
		f.SetCellValue(sheetName, gstCell, growthSTPct)

		// Apply style for Growth ST (%) based on value
		gstStyle := positiveGrowthStyle
		if data.GrowthSTPct < 0 {
			gstStyle = negativeGrowthStyle
		}
		f.SetCellStyle(sheetName, gstCell, gstCell, gstStyle)

	}

	// Adjust column widths
	f.SetColWidth(sheetName, "A", "A", 13) // Width for Dealer Code
	f.SetColWidth(sheetName, "B", "B", 46) // Width for Dealer Name
	f.SetColWidth(sheetName, "C", "C", 13) // Width for MTD SO
	f.SetColWidth(sheetName, "D", "D", 13) // Width for LMTD SO
	f.SetColWidth(sheetName, "E", "E", 13) // Width for Growth SO (%)
	f.SetColWidth(sheetName, "F", "F", 13) // Width for MTD SO
	f.SetColWidth(sheetName, "G", "G", 13) // Width for LMTD SO
	f.SetColWidth(sheetName, "H", "H", 13) // Width for Growth SO (%)
	f.SetColWidth(sheetName, "I", "I", 23) // Width for Growth SO (%)

	// Save the file
	if err := f.SaveAs(outputPath); err != nil {
		return fmt.Errorf("failed to save file: %w", err)
	}

	return nil
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

		dealerCode := row[5]
		tseName := row[15]

		if dealerCode != "" && tseName != "" {
			tseMapping[dealerCode] = tseName
		}
	}
	return tseMapping
}

// Function to run the entire growth report generation process
func RunGrowthReport() {
	dealerInfoFilePath := "../data/Dealer Information.xlsx"
	mtdSOFilePath := "../data/MTD-SO.xlsx"
	lmtdSOFilePath := "../data/LMTD-SO.xlsx"
	mtdSTFilePath := "../data/MTD-ST.xlsx"
	lmtdSTFilePath := "../data/LMTD-ST.xlsx"
	tseMappingFilePath := "../data/VIKING'S - DEALER Credit Period LIST.xlsx"

	// Get today's date for folder name
	today := time.Now().Format("2006-01-02")
	dirPath := fmt.Sprintf("./daily_growth_reports_%s", today)

	// Create directory for today's date
	if err := os.MkdirAll(dirPath, os.ModePerm); err != nil {
		log.Fatalf("Failed to create directory: %v", err)
	}

	outputPath := fmt.Sprintf("%s/daily_growth_report.xlsx", dirPath)

	dealerInfo, err := readDealerInformation(dealerInfoFilePath)
	if err != nil {
		log.Fatalf("Error reading dealer information: %v", err)
	}

	mtdSOData, err := readSellData(mtdSOFilePath, "SO")
	if err != nil {
		log.Fatalf("Error reading MTD data: %v", err)
	}

	lmtdSOData, err := readSellData(lmtdSOFilePath, "SO")
	if err != nil {
		log.Fatalf("Error reading LMTD data: %v", err)
	}

	mtdSTData, err := readSellData(mtdSTFilePath, "ST")
	if err != nil {
		log.Fatalf("Error reading MTD data: %v", err)
	}

	lmtdSTData, err := readSellData(lmtdSTFilePath, "ST")
	if err != nil {
		log.Fatalf("Error reading LMTD data: %v", err)
	}
	tseMapping := readTSEToRetailerMapping(tseMappingFilePath)
	report := generateGrowthReport(mtdSOData, lmtdSOData, mtdSTData, lmtdSTData, dealerInfo)

	if err := writeGrowthReport(report, tseMapping, outputPath); err != nil {
		log.Fatalf("Error writing growth report: %v", err)
	}

	fmt.Printf("Daily growth report generated successfully: %s\n", outputPath)
}
