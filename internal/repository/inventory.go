package repository

import (
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
	"viking-reports/internal/utils"

	"github.com/xuri/excelize/v2"
)

type ExcelInventoryRepository struct {
	filePath   string
	priceData  map[string]float64
	tseMapping map[string]string
}

type InventoryData struct {
	DealerCode         string
	DealerName         string
	TotalInventoryCost float64
	TotalCreditDue     float64
	TSE                string
	InventoryShortfall float64
}

func NewExcelInventoryRepository(filePath string, priceData map[string]float64, tseMapping map[string]string) *ExcelInventoryRepository {
	return &ExcelInventoryRepository{
		filePath:   filePath,
		priceData:  priceData,
		tseMapping: tseMapping,
	}
}

func (r *ExcelInventoryRepository) GetInventoryData() (map[string]*InventoryData, error) {
	f, err := excelize.OpenFile(r.filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open inventory file: %w", err)
	}
	defer f.Close()
	fmt.Println("Fetching today's stock inventory data for each retailer.")

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %w", err)
	}

	materialCodeIdx, err := utils.GetColumnIndex(f, sheetName, "Material Code")
	if err != nil {
		return nil, err
	}
	dealerCodeIdx, err := utils.GetColumnIndex(f, sheetName, "Dealer Code")
	if err != nil {
		return nil, err
	}
	dealerNameIdx, err := utils.GetColumnIndex(f, sheetName, "Dealer Name")
	if err != nil {
		return nil, err
	}

	inventoryData := make(map[string]*InventoryData)
	retailerCodeToCreditMap, _ := r.GetTotalCreditFromReports()
	for _, row := range rows[1:] {
		materialCode := row[materialCodeIdx]
		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]

		if materialCode == "" || dealerCode == "" {
			continue
		}

		netLandingCost := r.priceData[materialCode]
		if data, exists := inventoryData[dealerCode]; exists {
			data.TotalInventoryCost += netLandingCost
		} else {
			inventoryData[dealerCode] = &InventoryData{
				DealerCode:         dealerCode,
				DealerName:         dealerName,
				TSE:                r.tseMapping[dealerCode],
				TotalInventoryCost: netLandingCost,
			}
		}

	}

	// Update inventoryData with CostCreditDifference and TotalCredit
	for _, data := range inventoryData {
		data.TotalCreditDue = retailerCodeToCreditMap[data.DealerCode]
		data.InventoryShortfall = data.TotalInventoryCost - data.TotalCreditDue
	}

	return inventoryData, nil
}

func (r *ExcelInventoryRepository) GetTotalCreditFromReports() (map[string]float64, error) {
	creditData := make(map[string]float64)
	// Generate today's date in YYYY-MM-DD format
	today := time.Now().Format("2006-01-02")
	fmt.Printf("Fetching today's credit report for each retailer on %s\n", today) // Print key-value pairs
	reportDir := fmt.Sprintf("../credit_report/daily_credit_reports_%s", today)

	// Read all .xlsx files in the directory
	err := filepath.Walk(reportDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() && filepath.Ext(path) == ".xlsx" {
			f, err := excelize.OpenFile(path)
			if err != nil {
				return fmt.Errorf("failed to open credit report file %s: %w", path, err)
			}
			defer f.Close()

			sheetName := f.GetSheetName(0)
			rows, err := f.GetRows(sheetName)
			if err != nil {
				return fmt.Errorf("failed to get rows from %s: %w", path, err)
			}

			// Assuming the first row contains headers
			retailerCodeIdx, err := utils.GetColumnIndex(f, sheetName, "Retailer Code")
			if err != nil {
				return err
			}
			totalCreditIdx, err := utils.GetColumnIndex(f, sheetName, "Total Credit(₹)")
			if err != nil {
				return err
			}

			for i, row := range rows[1:] {
				if i == len(rows)-2 { // Skip the last row
					break
				}
				if len(row) <= totalCreditIdx {
					continue
				}
				retailerCode := row[retailerCodeIdx]
				totalCredit := row[totalCreditIdx]
				totalCredit = strings.ReplaceAll(totalCredit, ",", "") // Remove commas

				val, err := strconv.ParseFloat(totalCredit, 64) // Convert string to float64
				if err != nil {
					return fmt.Errorf("failed to parse total credit for retailer %s: %w", retailerCode, err)
				}
				creditData[retailerCode] += val // Accumulate total credit for ea
			}
		}

		return nil
	})

	if err != nil {
		return nil, err
	}

	return creditData, nil
}
