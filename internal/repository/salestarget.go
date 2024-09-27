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
	ItemName   string
}

type ExcelSalesTargetRepository struct {
}

func NewExcelSalesTargetRepository() *ExcelSalesTargetRepository {
	return &ExcelSalesTargetRepository{}
}

func (r *ExcelSalesTargetRepository) ReadSales(salesFilePath string, tseMap map[string]string) ([]*SalesData, error) {
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

	itemNameIdx, err := utils.GetHeaderIndex(f, sheetName, "Item Name", 9)
	if err != nil {
		return nil, err
	}

	sales := make([]*SalesData, 0)
	for _, row := range rows[9:] {
		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]
		itemName := row[itemNameIdx]
		if dealerCode == "" {
			continue
		}

		amountStr := row[saleAmountIdx]
		amount, err := strconv.ParseFloat(amountStr, 64) // Convert to int
		if err != nil {
			continue // Ignore row if conversion fails
		}

		// Create SalesData object and add to slice
		sales = append(sales, &SalesData{
			DealerCode: dealerCode,
			DealerName: dealerName,
			MTDS:       1, // Set MTDS to 1 for each entry
			Value:      int(amount),
			TSE:        tseMap[dealerCode],
			ItemName:   itemName,
		})

	}

	return sales, nil
}
