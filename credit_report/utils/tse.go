package utils

import (
	"log"

	"github.com/xuri/excelize/v2"
)

// Read TSE to retailer mapping from an Excel file
func ReadTSEToRetailerMapping(tseMappingFilePath string) map[string]string {
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

// Create a map of retailer name to code from the TSE mapping file
func CreateRetailerNameToCodeMap(tseMapFilePath string) map[string]string {
	tseMappingFile, err := excelize.OpenFile(tseMapFilePath)
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

		dealerCode, dealerName := row[5], row[6]
		if dealerCode != "" && dealerName != "" {
			tseMapping[dealerName] = dealerCode
		}
	}
	return tseMapping
}
