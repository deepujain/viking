package repository

import (
	"fmt"
	"strings"
	"time"
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
	GrowthSOPct int
	MTDST       int
	LMTDST      int
	GrowthSTPct int
}

type SellData struct {
	DealerCode string
	DealerName string
	MTDS       int
	Date       string
}

type DealerSPUSales struct {
	DealerCode string
	DealerName string
	SPUName    string
	Count      int
}

func NewExcelSalesRepository() *ExcelSalesRepository {
	return &ExcelSalesRepository{}
}

func (r *ExcelSalesRepository) GetSales(salesFilePath string) (map[string]*SellData, error) {
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
	activateTimeIdx, err := utils.GetColumnIndex(f, sheetName, "Activate Time")
	if err != nil {
		// Try with "activateTime" if "Activate Time" is not found
		activateTimeIdx, err = utils.GetColumnIndex(f, sheetName, "activateTime")
		if err != nil {
			return nil, err
		}
	}

	sellData := make(map[string]*SellData)
	for _, row := range rows[1:] {
		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]
		activationTime := row[activateTimeIdx]

		// Filter lmtdSOData to keep only records with Date less than today
		today := time.Now()
		todayDay := today.Day()
		saleDate, _ := time.Parse("2006-01-02 15:04:05", activationTime)
		saleDay := saleDate.Day() // Extract day from activation time

		if dealerCode == "" || saleDay > todayDay {
			if dealerCode == "IN001525" {
				fmt.Println(saleDate)
			}
			continue
		}

		if data, exists := sellData[dealerCode]; exists {
			data.MTDS++
		} else {
			sellData[dealerCode] = &SellData{
				DealerCode: dealerCode,
				DealerName: dealerName,
				Date:       activationTime,
				MTDS:       1,
			}
		}

	}

	return sellData, nil
}

func (r *ExcelSalesRepository) GetDealerSPUSales(salesFilePath string) (map[string]*DealerSPUSales, error) {
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

	spuNameIdx, err := utils.GetColumnIndex(f, sheetName, "SPU Name")
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
	productTypeIdx, err := utils.GetColumnIndex(f, sheetName, "Product Type")
	if err != nil {
		return nil, err
	}

	dealerSPUSales := make(map[string]*DealerSPUSales)
	for _, row := range rows[1:] {
		rawSpuName := row[spuNameIdx]
		spuName := strings.ReplaceAll(rawSpuName, "realme", "") // Remove "realme" from model
		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]
		productType := row[productTypeIdx]

		if spuName == "" || dealerCode == "" || dealerName == "" || !strings.Contains(productType, "mobile") {
			continue
		}
		// {{ edit_1 }}: Calculate quantity (QTY) for each retailer and SPU Name
		if data, exists := dealerSPUSales[dealerName+spuName]; exists {
			data.Count += 1 // Increment count for existing SPU
		} else {
			dealerSPUSales[dealerName+spuName] = &DealerSPUSales{
				DealerCode: dealerCode,
				DealerName: dealerName,
				SPUName:    spuName,
				Count:      1, // Initialize count
			}
		}
	}
	return dealerSPUSales, nil
}
