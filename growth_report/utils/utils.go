// Package utils provides utility functions for generating growth reports.
package utils

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sort"
	"time"

	"github.com/xuri/excelize/v2"
)

// SellData represents the sell data for each dealer.
type SellData struct {
	DealerCode string
	DealerName string
	MTDS       int
	LMTDS      int
}

// GrowthReport represents the final growth report data for each dealer.
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

// readDealerInformation reads dealer information from an Excel file.
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

	dealerCodeIdx, err := getColumnIndex(xlFile, sheetName, "Dealer Code")
	if err != nil {
		return nil, fmt.Errorf("failed to get Dealer Code column index: %w", err)
	}
	dealerNameIdx, err := getColumnIndex(xlFile, sheetName, "Dealer Name")
	if err != nil {
		return nil, fmt.Errorf("failed to get Dealer Name column index: %w", err)
	}

	dealerInfo := make(map[string]string)
	for i, row := range rows[1:] { // Skip header row
		if dealerCode := row[dealerCodeIdx]; dealerCode != "" {
			dealerInfo[dealerCode] = row[dealerNameIdx]
		} else {
			log.Printf("Skipping row %d due to empty Dealer Code", i+2)
		}
	}

	return dealerInfo, nil
}

// readSellData reads sell-out or sell-through data from an Excel file.
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
	currentDay := time.Now().Day()

	for _, row := range rows[1:] { // Skip header row
		activateTime, err := time.Parse("2006-01-02 15:04:05", row[activateTimeIdx])
		if err != nil || activateTime.Day() > currentDay {
			continue
		}

		dealerCode := row[dealerCodeIdx]
		if dealerCode == "" {
			continue
		}

		if data, exists := sellData[dealerCode]; exists {
			data.MTDS++
		} else {
			sellData[dealerCode] = &SellData{
				DealerCode: dealerCode,
				DealerName: row[dealerNameIdx],
				MTDS:       1,
			}
		}
	}

	return sellData, nil
}

// generateGrowthReport generates the growth report based on sell data.
func generateGrowthReport(mtdSOData, lmtdSOData, mtdSTData, lmtdSTData map[string]*SellData, dealerInfo map[string]string) []GrowthReport {
	var report []GrowthReport

	for dealerCode, dealerName := range dealerInfo {
		mtdSO := getOrCreateSellData(mtdSOData, dealerCode, dealerName)
		lmtdSO := getOrCreateSellData(lmtdSOData, dealerCode, dealerName)
		mtdST := getOrCreateSellData(mtdSTData, dealerCode, dealerName)
		lmtdST := getOrCreateSellData(lmtdSTData, dealerCode, dealerName)

		reportEntry := GrowthReport{
			DealerCode:  dealerCode,
			DealerName:  dealerName,
			MTDSO:       mtdSO.MTDS,
			LMTDSO:      lmtdSO.MTDS,
			GrowthSOPct: calculateGrowthPercentage(mtdSO.MTDS, lmtdSO.MTDS),
			MTDST:       mtdST.MTDS,
			LMTDST:      lmtdST.MTDS,
			GrowthSTPct: calculateGrowthPercentage(mtdST.MTDS, lmtdST.MTDS),
		}
		report = append(report, reportEntry)
	}

	sort.Slice(report, func(i, j int) bool {
		return report[j].GrowthSOPct > report[i].GrowthSOPct
	})

	return report
}

// getOrCreateSellData retrieves or creates a SellData object.
func getOrCreateSellData(data map[string]*SellData, dealerCode, dealerName string) *SellData {
	if data, exists := data[dealerCode]; exists {
		return data
	}
	return &SellData{DealerCode: dealerCode, DealerName: dealerName, MTDS: 0, LMTDS: 0}
}

// calculateGrowthPercentage calculates the growth percentage.
func calculateGrowthPercentage(current, previous int) float64 {
	if previous > 0 {
		return float64(current-previous) / float64(previous) * 100
	}
	if current > 0 {
		return 100.0
	}
	return 0.0
}

// writeGrowthReport writes the growth report to an Excel file.
func writeGrowthReport(report []GrowthReport, tseMapping map[string]string, outputPath string) error {
	f := excelize.NewFile()
	sheetName := "Growth Report"
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to create new sheet: %w", err)
	}
	f.SetActiveSheet(index)

	if err := writeHeaders(f, sheetName); err != nil {
		return err
	}

	if err := writeData(f, sheetName, report, tseMapping); err != nil {
		return err
	}

	adjustColumnWidths(f, sheetName)

	if err := f.SaveAs(outputPath); err != nil {
		return fmt.Errorf("failed to save file: %w", err)
	}

	return nil
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

	return nil
}

// writeData writes the report data to the Excel sheet.
func writeData(f *excelize.File, sheetName string, report []GrowthReport, tseMapping map[string]string) error {
	positiveGrowthStyle, negativeGrowthStyle, err := createGrowthStyles(f)
	if err != nil {
		return err
	}

	for row, data := range report {
		rowIndex := row + 2
		if err := writeSingleRow(f, sheetName, rowIndex, data, tseMapping, positiveGrowthStyle, negativeGrowthStyle); err != nil {
			return err
		}
	}

	return nil
}

// createGrowthStyles creates styles for positive and negative growth.
func createGrowthStyles(f *excelize.File) (int, int, error) {
	positiveGrowthStyle, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "000000"},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"00FF00"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return 0, 0, fmt.Errorf("failed to create positive growth style: %w", err)
	}

	negativeGrowthStyle, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "000000"},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"FF0000"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return 0, 0, fmt.Errorf("failed to create negative growth style: %w", err)
	}

	return positiveGrowthStyle, negativeGrowthStyle, nil
}

// writeSingleRow writes a single row of data to the Excel sheet.
func writeSingleRow(f *excelize.File, sheetName string, rowIndex int, data GrowthReport, tseMapping map[string]string, positiveGrowthStyle, negativeGrowthStyle int) error {
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), data.DealerCode)
	f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIndex), data.DealerName)
	f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowIndex), data.MTDSO)
	f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowIndex), data.LMTDSO)
	f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowIndex), int(data.GrowthSOPct))
	f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowIndex), data.MTDST)
	f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowIndex), data.LMTDST)
	f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowIndex), int(data.GrowthSTPct))
	f.SetCellValue(sheetName, fmt.Sprintf("I%d", rowIndex), tseMapping[data.DealerCode])

	// Set borders for all cells in the row
	cells := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I"}
	borderStyle := excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	}
	borderStyleID, err := f.NewStyle(&borderStyle)
	if err != nil {
		return err
	}
	for _, cell := range cells {
		cellRef := fmt.Sprintf("%s%d", cell, rowIndex)
		f.SetCellStyle(sheetName, cellRef, cellRef, borderStyleID)
	}

	// Apply conditional styles for growth percentages
	gsoCell := fmt.Sprintf("E%d", rowIndex)
	gstCell := fmt.Sprintf("H%d", rowIndex)

	gsoStyle := positiveGrowthStyle
	if data.GrowthSOPct < 0 {
		gsoStyle = negativeGrowthStyle
	}
	f.SetCellStyle(sheetName, gsoCell, gsoCell, gsoStyle)

	gstStyle := positiveGrowthStyle
	if data.GrowthSTPct < 0 {
		gstStyle = negativeGrowthStyle
	}
	f.SetCellStyle(sheetName, gstCell, gstCell, gstStyle)

	return nil
}

// adjustColumnWidths adjusts the widths of columns in the Excel sheet.
func adjustColumnWidths(f *excelize.File, sheetName string) {
	columnWidths := map[string]float64{
		"A": 13, // Dealer Code
		"B": 46, // Dealer Name
		"C": 13, // MTD SO
		"D": 13, // LMTD SO
		"E": 15, // Growth SO (%)
		"F": 13, // MTD ST
		"G": 13, // LMTD ST
		"H": 13, // Growth ST (%)
		"I": 19, // TSE
	}

	for col, width := range columnWidths {
		f.SetColWidth(sheetName, col, col, width)
	}
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

// RunGrowthReport runs the entire growth report generation process.
func RunGrowthReport() error {
	dataDir := "../data"
	files := map[string]string{
		"dealerInfo": "Dealer Information.xlsx",
		"mtdSO":      "MTD-SO.xlsx",
		"lmtdSO":     "LMTD-SO.xlsx",
		"mtdST":      "MTD-ST.xlsx",
		"lmtdST":     "LMTD-ST.xlsx",
		"tseMapping": "VIKING'S - DEALER Credit Period LIST.xlsx",
	}

	for key, filename := range files {
		files[key] = filepath.Join(dataDir, filename)
	}

	today := time.Now().Format("2006-01-02")
	dirPath := fmt.Sprintf("./daily_growth_reports_%s", today)
	if err := os.MkdirAll(dirPath, os.ModePerm); err != nil {
		return fmt.Errorf("failed to create directory: %w", err)
	}

	outputPath := filepath.Join(dirPath, "daily_growth_report.xlsx")

	dealerInfo, err := readDealerInformation(files["dealerInfo"])
	if err != nil {
		return fmt.Errorf("error reading dealer information: %w", err)
	}

	mtdSOData, err := readSellData(files["mtdSO"], "SO")
	if err != nil {
		return fmt.Errorf("error reading MTD SO data: %w", err)
	}

	lmtdSOData, err := readSellData(files["lmtdSO"], "SO")
	if err != nil {
		return fmt.Errorf("error reading LMTD SO data: %w", err)
	}

	mtdSTData, err := readSellData(files["mtdST"], "ST")
	if err != nil {
		return fmt.Errorf("error reading MTD ST data: %w", err)
	}

	lmtdSTData, err := readSellData(files["lmtdST"], "ST")
	if err != nil {
		return fmt.Errorf("error reading LMTD ST data: %w", err)
	}

	tseMapping, err := readTSEToRetailerMapping(files["tseMapping"])
	if err != nil {
		return fmt.Errorf("error reading TSE mapping: %w", err)
	}

	report := generateGrowthReport(mtdSOData, lmtdSOData, mtdSTData, lmtdSTData, dealerInfo)

	if err := writeGrowthReport(report, tseMapping, outputPath); err != nil {
		return fmt.Errorf("error writing growth report: %w", err)
	}

	fmt.Printf("Daily growth report generated successfully: %s\n", outputPath)
	return nil
}
