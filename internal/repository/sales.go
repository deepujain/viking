package repository

import (
	"fmt"
	"viking-reports/internal/utils"

	"github.com/xuri/excelize/v2"
)

type ExcelSalesRepository struct {
	filePath string
}

type SellData struct {
	DealerCode string
	DealerName string
	MTDS       int
}

func NewExcelSalesRepository(filePath string) *ExcelSalesRepository {
	return &ExcelSalesRepository{filePath: filePath}
}

func (r *ExcelSalesRepository) GetSellData(fileType string) (map[string]*SellData, error) {
	f, err := excelize.OpenFile(r.filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open sales file: %w", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %w", err)
	}

	dealerCodeIdx, err := utils.GetColumnIndex(f, sheetName, "Dealer Code")
	if err != nil {
		return nil, err
	}
	dealerNameIdx, err := utils.GetColumnIndex(f, sheetName, "Dealer Name")
	if err != nil {
		return nil, err
	}

	sellData := make(map[string]*SellData)
	for _, row := range rows[1:] {
		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]

		if dealerCode == "" {
			continue
		}

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
