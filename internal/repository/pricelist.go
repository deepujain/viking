package repository

import (
	"strconv"

	"github.com/xuri/excelize/v2"
)

type ExcelPriceListRepository struct {
	filePath string
}

type PriceListData struct {
	RetailerCode string
	REALMEMobile string
}

type PriceListRow struct {
	Type     string
	Model    string
	Color    string
	Capacity string
	NLC      int
	Mop      int
	Mrp      int
}

func NewExcelPriceListRepository(filePath string) *ExcelPriceListRepository {
	return &ExcelPriceListRepository{filePath: filePath}
}

func (r *ExcelPriceListRepository) GetPriceListData() ([]PriceListRow, error) {
	f, err := excelize.OpenFile(r.filePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	var results []PriceListRow
	sheetName := f.GetSheetName(0) // Assuming the first sheet is the one we need
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	var lastType, lastModel, lastColor string // Track last seen values

	for i, row := range rows {
		if i == 0 || i == 1 {
			// Skip header
			continue
		}
		if len(row) >= 7 { // Ensure we have enough columns
			// Check for merged values
			if row[0] != "" {
				lastType = row[0]
			}
			if row[1] != "" {
				lastModel = row[1]
			}
			if row[2] != "" {
				lastColor = row[2]
			}

			// Use last seen values if current row is missing
			model := lastModel
			color := lastColor
			capacity := row[3]
			dlrPrice, _ := strconv.Atoi(row[4])
			mop, _ := strconv.Atoi(row[5])
			mrp, _ := strconv.Atoi(row[6])

			// Create a new PriceListRow struct and append to results
			priceListRow := PriceListRow{
				Type:     lastType,
				Model:    model,
				Color:    color,
				Capacity: capacity,
				NLC:      dlrPrice,
				Mop:      mop,
				Mrp:      mrp,
			}
			results = append(results, priceListRow)
		}
	}

	// Return the results array
	return results, nil
}
