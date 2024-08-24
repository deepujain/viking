package repository

import (
	"fmt"
	"viking-reports/internal/utils"

	"github.com/xuri/excelize/v2"
)

type ExcelSalesRepository struct {
}
type GrowthData struct {
	DealerCode  string
	DealerName  string
	MTDSO       int
	LMTDSO      int
	GrowthSOPct float64
	MTDST       int
	LMTDST      int
	GrowthSTPct float64
}

type SellData struct {
	DealerCode string
	DealerName string
	MTDS       int
}

func NewExcelSalesRepository() *ExcelSalesRepository {
	return &ExcelSalesRepository{}
}

func (r *ExcelSalesRepository) GetSellData(salesFilePath string) (map[string]*SellData, error) {
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

	dealerCodeIdx, err := utils.GetColumnIndex(f, sheetName, "Dealer Code")
	if err != nil {
		// Try with "toDealerCode" if "Dealer Code" is not found
		dealerCodeIdx, err = utils.GetColumnIndex(f, sheetName, "toDealerCode")
		if err != nil {
			return nil, err
		}
	}
	dealerNameIdx, err := utils.GetColumnIndex(f, sheetName, "Dealer Name")
	if err != nil {
		// Try with "toDealerName" if "Dealer Name" is not found
		dealerNameIdx, err = utils.GetColumnIndex(f, sheetName, "toDealerName")
		if err != nil {
			return nil, err
		}
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
