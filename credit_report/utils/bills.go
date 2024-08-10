package utils

import (
	"log"
	"strconv"

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

// Read bills from an Excel file
func ReadBills(inputFilePath string) []Bill {
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
