package repository

import (
	"fmt"
	"strconv"
	"viking-reports/internal/utils"

	"github.com/xuri/excelize/v2"
)

type SalesData struct {
	DealerCode string
	DealerName string
	MTDS       int
	TSE        string
	Value      int
}

type ExcelSalesTargetRepository struct {
}

func NewExcelSalesTargetRepository() *ExcelSalesTargetRepository {
	return &ExcelSalesTargetRepository{}
}

func (r *ExcelSalesTargetRepository) ComputeSales(salesFilePath string, tseMap map[string]string) (map[string]*SalesData, error) {
	fmt.Printf(" from %s \n", salesFilePath)
	f, err := excelize.OpenFile(salesFilePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open sales file: %w", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %w", err)
	}
	dealerCodeIdx, err := utils.GetHeaderIndex(f, sheetName, "Retailer Code", 9)
	if err != nil {

		return nil, err

	}
	dealerNameIdx, err := utils.GetHeaderIndex(f, sheetName, "Party Name", 9)
	if err != nil {
		return nil, err
	}
	saleAmountIdx, err := utils.GetHeaderIndex(f, sheetName, "Amount ", 9)
	if err != nil {
		return nil, err
	}

	sales := make(map[string]*SalesData)
	for _, row := range rows[9:] {
		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]

		if dealerCode == "" {
			continue
		}

		if data, exists := sales[dealerCode]; exists {
			data.MTDS++
			amountStr := row[saleAmountIdx]
			amount, err := strconv.ParseFloat(amountStr, 64) // Convert to int
			if err != nil {
				return nil, fmt.Errorf("failed to convert sale amount to int: %w", err)
			}
			data.Value += int(amount)
		} else {
			sales[dealerCode] = &SalesData{
				DealerCode: dealerCode,
				DealerName: dealerName,
				MTDS:       1,
				Value:      0,
				TSE:        tseMap[dealerCode],
			}
		}

	}

	return sales, nil
}
