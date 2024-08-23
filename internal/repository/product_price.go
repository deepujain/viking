package repository

import (
	"fmt"
	"strconv"
	"viking-reports/internal/utils"

	"github.com/xuri/excelize/v2"
)

type ExcelProductPriceRepository struct {
	filePath string
}

func NewExcelProductPriceRepository(filePath string) *ExcelProductPriceRepository {
	return &ExcelProductPriceRepository{filePath: filePath}
}

func (r *ExcelProductPriceRepository) GetProductPrices() (map[string]float64, error) {
	f, err := excelize.OpenFile(r.filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open price list file: %w", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %w", err)
	}

	productCodeIdx, err := utils.GetColumnIndex(f, sheetName, "Material Code")
	if err != nil {
		return nil, err
	}
	priceIdx, err := utils.GetColumnIndex(f, sheetName, "NLC")
	if err != nil {
		return nil, err
	}

	priceData := make(map[string]float64)
	for _, row := range rows[1:] {
		productCode := row[productCodeIdx]
		priceStr := row[priceIdx]

		if productCode == "" {
			continue
		}

		price, err := strconv.ParseFloat(priceStr, 64)
		if err != nil {
			return nil, fmt.Errorf("failed to parse price for product %s: %w", productCode, err)
		}

		priceData[productCode] = price
	}

	return priceData, nil
}
