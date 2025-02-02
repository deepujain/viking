package repository

import (
	"fmt"
	"log"
	"path/filepath"
	"strconv"
	"time"
	"viking-reports/internal/utils"

	"github.com/xuri/excelize/v2"
)

type ExcelCreditRepository struct {
	filePath string
}

type CreditData struct {
	RetailerCode string
	TotalCredit  float64
}

type Bill struct {
	Date          string
	RefNo         string
	RetailerName  string
	PendingAmount float64
	DueDate       string
	AgeOfBill     int

	TSE string
}

func NewExcelCreditRepository(filePath string) *ExcelCreditRepository {
	return &ExcelCreditRepository{filePath: filePath}
}

func (r *ExcelCreditRepository) GetCreditData() (map[string]*CreditData, error) {
	today := time.Now().Format("2006-01-02")
	creditData := make(map[string]*CreditData)

	files, err := filepath.Glob(filepath.Join(r.filePath, fmt.Sprintf("credit_reports_%s", today), "*.xlsx"))
	if err != nil {
		return nil, fmt.Errorf("failed to find credit report files: %w", err)
	}

	for _, file := range files {
		f, err := excelize.OpenFile(file)
		if err != nil {
			return nil, fmt.Errorf("failed to open credit file %s: %w", file, err)
		}
		defer f.Close()

		sheetName := f.GetSheetName(0)
		rows, err := f.GetRows(sheetName)
		if err != nil {
			return nil, fmt.Errorf("failed to get rows from %s: %w", file, err)
		}

		retailerCodeIdx, err := utils.GetColumnIndex(f, sheetName, "Retailer Code")
		if err != nil {
			return nil, err
		}
		totalCreditIdx, err := utils.GetColumnIndex(f, sheetName, "Total Credit")
		if err != nil {
			return nil, err
		}

		for _, row := range rows[1:] {
			retailerCode := row[retailerCodeIdx]
			totalCredit := utils.ParseFloat(row[totalCreditIdx])

			if data, exists := creditData[retailerCode]; exists {
				data.TotalCredit += totalCredit
			} else {
				creditData[retailerCode] = &CreditData{
					RetailerCode: retailerCode,
					TotalCredit:  totalCredit,
				}
			}
		}
	}

	return creditData, nil
}

func (r *ExcelCreditRepository) AggregateCreditByRetailer(bills []Bill, tseMapping map[string]string, retailerNameToCodeMap map[string]string) map[string]map[string]interface{} {
	aggregatedData := make(map[string]map[string]interface{})

	// Step 1: Group bills by retailer name
	groupedBills := make(map[string][]Bill)
	for _, bill := range bills {
		groupedBills[bill.RetailerName] = append(groupedBills[bill.RetailerName], bill)
	}

	// Step 2: Process each group of bills
	for retailerName, retailerBills := range groupedBills {
		totalPendingAmount := 0.0
		// Initialize retailer data if it doesn't exist
		if _, exists := aggregatedData[retailerName]; !exists {
			aggregatedData[retailerName] = make(map[string]interface{})
			aggregatedData[retailerName]["Retailer Code"] = retailerNameToCodeMap[retailerName]
			aggregatedData[retailerName]["Retailer Name"] = retailerName
			aggregatedData[retailerName]["0-7 Days"] = 0.0
			aggregatedData[retailerName]["8-14 Days"] = 0.0
			aggregatedData[retailerName]["15-20 Days"] = 0.0
			aggregatedData[retailerName]["21-30 Days"] = 0.0
			aggregatedData[retailerName]["31+ Days"] = 0.0
			aggregatedData[retailerName]["Total Credit"] = 0.0
			aggregatedData[retailerName]["TSE"] = tseMapping[retailerName]
		}

		// Step 3: Calculate total pending amount and update days based on age of bill
		for _, bill := range retailerBills {
			totalPendingAmount += bill.PendingAmount

			// Update the days based on the age of the bill
			switch {
			case bill.AgeOfBill >= 0 && bill.AgeOfBill <= 7:
				aggregatedData[retailerName]["0-7 Days"] = aggregatedData[retailerName]["0-7 Days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill >= 8 && bill.AgeOfBill <= 14:
				aggregatedData[retailerName]["8-14 Days"] = aggregatedData[retailerName]["8-14 Days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill >= 15 && bill.AgeOfBill <= 20:
				aggregatedData[retailerName]["15-20 Days"] = aggregatedData[retailerName]["15-20 Days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill >= 21 && bill.AgeOfBill <= 30:
				aggregatedData[retailerName]["21-30 Days"] = aggregatedData[retailerName]["21-30 Days"].(float64) + bill.PendingAmount
			default:
				aggregatedData[retailerName]["31+ Days"] = aggregatedData[retailerName]["31+ Days"].(float64) + bill.PendingAmount
			}
		}

		// Step 4: Update total credit for this retailer
		aggregatedData[retailerName]["Total Credit"] = totalPendingAmount
	}
	return aggregatedData
}

func (r *ExcelCreditRepository) GetBills() ([]Bill, error) {
	fmt.Println("** Input: Fetching invoices of daily sales from Tally. **")

	f, err := excelize.OpenFile(r.filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open bills file: %w", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %w", err)
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

	return bills, nil
}
