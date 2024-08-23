package repository

import (
	"fmt"
	"viking-reports/internal/utils"

	"github.com/xuri/excelize/v2"
)

type ExcelInventoryRepository struct {
	filePath   string
	priceData  map[string]float64
	tseMapping map[string]string
}

type InventoryData struct {
	DealerCode           string
	DealerName           string
	TotalInventoryCost   float64
	TotalCreditDue       float64
	TSE                  string
	CostCreditDifference float64
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

		// Update inventoryData with CostCreditDifference and TotalCredit
		for _, data := range inventoryData {
			data.TotalCreditDue = 100
			data.CostCreditDifference = data.TotalInventoryCost - data.TotalCreditDue
		}
	}

	return inventoryData, nil
}
