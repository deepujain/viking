package main

import (
	"fmt"
	"log"
	"os"
	"time"

	"credit_report/utils"
)

func main() {
	inputFilePath := "../data/credit_report/Bills.xlsx"                              // Read from current directory
	tseMappingFilePath := "../data/common/VIKING'S - DEALER Credit Period LIST.xlsx" // Read the TSE mapping file

	// Get today's date for folder name
	today := time.Now().Format("2006-01-02")
	dirPath := fmt.Sprintf("./daily_credit_reports_%s", today)

	// Create directory for today's date
	if err := os.MkdirAll(dirPath, os.ModePerm); err != nil {
		log.Fatalf("Failed to create directory: %v", err)
	}

	bills := utils.ReadBills(inputFilePath)

	retailerNameToTSEMap := utils.ReadTSEToRetailerMapping(tseMappingFilePath)
	retailerNameToCodeMap := utils.CreateRetailerNameToCodeMap(tseMappingFilePath)

	// Categorize bills by retailer and age category
	aggregatedData := utils.AggregateCreditByRetailer(bills, retailerNameToTSEMap)

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
		if err := utils.WriteCategorizedData(dirPath, fileName, data, retailerNameToCodeMap); err != nil {
			log.Printf("Error writing file for TSE %s: %v", tseName, err)
		} else {
			log.Printf("Credit report for TSE %s saved to: %s/%s\n", tseName, dirPath, fileName)
		}
	}

	// Write output file for missing TSE
	if len(missingTSEData) > 0 {
		if err := utils.WriteCategorizedData(dirPath, "TSE_MISSING_credit_report.xlsx", missingTSEData, retailerNameToCodeMap); err != nil {
			log.Printf("Error writing TSE_MISSING file: %v", err)
		} else {
			log.Printf("Credit report for missing TSE saved to: %s/TSE_MISSING_credit_report.xlsx\n", dirPath)
		}
	}
}
