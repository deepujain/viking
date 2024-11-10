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

type InventoryShortFallRepo struct {
	DealerCode         string
	DealerName         string
	TotalInventoryCost float64
	TotalCreditDue     float64
	TSE                string
	InventoryShortfall float64
}

type ModelCountRepo struct {
	DealerCode   string
	DealerName   string
	MaterialCode int
	SPUName      string
	Color        string
	SKUSpec      string
	ProductType  string
	TSE          string
	Count        int
}

type SPUInventoryCount struct {
	DealerCode string
	DealerName string
	SPUName    string
	Count      int
}

func NewSPUInventoryRepository(filePath string) *ExcelInventoryRepository {
	return &ExcelInventoryRepository{
		filePath: filePath,
	}
}

func NewExcelInventoryRepository(filePath string, priceData map[string]float64, tseMapping map[string]string) *ExcelInventoryRepository {
	return &ExcelInventoryRepository{
		filePath:   filePath,
		priceData:  priceData,
		tseMapping: tseMapping,
	}
}

func (r *ExcelInventoryRepository) ComputeDealerSPUInventory() (map[string]*SPUInventoryCount, error) {
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

	dealerSPUInventory := make(map[string]*SPUInventoryCount)
	for _, row := range rows[1:] {
		rawSpuName := row[spuNameIdx]
		spuName := strings.ReplaceAll(rawSpuName, "realme", "") // Remove "realme" from model
		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]

		if spuName == "" || dealerCode == "" || dealerName == "" {
			continue
		}
		// {{ edit_1 }}: Calculate quantity (QTY) for each retailer and SPU Name
		if data, exists := dealerSPUInventory[dealerName+spuName]; exists {
			data.Count += 1 // Increment count for existing SPU
		} else {
			dealerSPUInventory[dealerName+spuName] = &SPUInventoryCount{
				DealerCode: dealerCode,
				DealerName: dealerName,
				SPUName:    spuName,
				Count:      1, // Initialize count
			}
		}
	}
	return dealerSPUInventory, nil
}

func (r *ExcelInventoryRepository) ComputeInventoryShortFall() (map[string]*InventoryShortFallRepo, error) {
	fmt.Println("Compute current inventory and shortfall for all retailers.")
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

	inventoryData := make(map[string]*InventoryShortFallRepo)
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
			inventoryData[dealerCode] = &InventoryShortFallRepo{
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
	reportDir := fmt.Sprintf("../credit_report/credit_reports_%s", today)

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

func (r *ExcelInventoryRepository) ComputeMaterialModelCount() (map[string]*ModelCountRepo, error) {
	fmt.Printf("Computing material model count for all retailers using %s \n", r.filePath)
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
	spuNameIdx, err := utils.GetColumnIndex(f, sheetName, "SPU Name")
	if err != nil {
		return nil, err
	}
	colorIdx, err := utils.GetColumnIndex(f, sheetName, "Color")
	if err != nil {
		return nil, err
	}
	skuSpecIdx, err := utils.GetColumnIndex(f, sheetName, "SKU Spec")
	if err != nil {
		return nil, err
	}
	productTypeIdx, err := utils.GetColumnIndex(f, sheetName, "Product Type")
	if err != nil {
		return nil, err
	}

	distributorNameIdx, err := utils.GetColumnIndex(f, sheetName, "Area Name")
	if err != nil {
		return nil, err
	}

	materialCount := make(map[string]*ModelCountRepo)
	for _, row := range rows[1:] {
		materialCode := row[materialCodeIdx]
		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]

		if materialCode == "" {
			continue
		}

		if dealerCode == "" {
			dealerName = row[distributorNameIdx]
		}

		if data, exists := materialCount[materialCode]; exists {
			data.Count += 1 // Increment count for existing dealer
		} else {
			materialCodeInt, _ := strconv.Atoi(materialCode)
			materialCount[materialCode] = &ModelCountRepo{
				DealerCode:   dealerCode,
				DealerName:   dealerName,
				MaterialCode: materialCodeInt,
				SPUName:      row[spuNameIdx],
				Color:        row[colorIdx],
				SKUSpec:      row[skuSpecIdx],
				ProductType:  row[productTypeIdx],
				Count:        1, // Initialize count
				TSE:          r.tseMapping[dealerCode],
			}
		}
	}

	return materialCount, nil
}
