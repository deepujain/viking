package repository

import (
	"fmt"
	"strconv"
	"viking-reports/internal/utils"

	"github.com/xuri/excelize/v2"
)

type ExcelTSEMappingRepository struct {
	filePath string
}

func NewExcelTSEMappingRepository(filePath string) *ExcelTSEMappingRepository {
	return &ExcelTSEMappingRepository{filePath: filePath}
}

func (r *ExcelTSEMappingRepository) GetRetailerCodeToTSEMap() (map[string]string, error) {
	fmt.Println("Input: Fetching retailer code to TSE name map from ", r.filePath)

	f, err := excelize.OpenFile(r.filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open TSE mapping file: %w", err)
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
	tseNameIdx, err := utils.GetColumnIndex(f, sheetName, "TSE Name")
	if err != nil {
		return nil, err
	}

	tseMapping := make(map[string]string)
	for _, row := range rows[1:] {
		dealerCode := row[dealerCodeIdx]
		tseName := row[tseNameIdx]

		if dealerCode == "" {
			continue
		}
		tseMapping[dealerCode] = tseName
	}

	return tseMapping, nil
}

func (r *ExcelTSEMappingRepository) GetRetailerCodeToNameMap() (map[string]string, error) {
	fmt.Println("Input: Fetching retailer code to TSE name map from ", r.filePath)

	f, err := excelize.OpenFile(r.filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open TSE mapping file: %w", err)
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

	codeToDealerMap := make(map[string]string)
	for _, row := range rows[1:] {
		dealerCode := row[dealerCodeIdx]
		dealerName := row[dealerNameIdx]

		if dealerCode == "" {
			continue
		}
		codeToDealerMap[dealerCode] = dealerName
	}

	return codeToDealerMap, nil
}

func (r *ExcelTSEMappingRepository) GetRetailerNameToTSEMap(dealerNameHeader string) (map[string]string, error) {
	fmt.Println("Input: Fetching retailer name to TSE name map from metadata ", r.filePath)
	f, err := excelize.OpenFile(r.filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open TSE mapping file: %w", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %w", err)
	}

	dealerNameIdx, err := utils.GetColumnIndex(f, sheetName, dealerNameHeader)
	if err != nil {
		return nil, err
	}
	tseNameIdx, err := utils.GetColumnIndex(f, sheetName, "TSE Name")
	if err != nil {
		return nil, err
	}

	tseMapping := make(map[string]string)
	for _, row := range rows[1:] {
		dealerName := row[dealerNameIdx]
		tseName := row[tseNameIdx]

		if dealerName == "" {
			continue
		}
		tseMapping[dealerName] = tseName
	}

	return tseMapping, nil
}

func (r *ExcelTSEMappingRepository) GetRetailerNameToCodeMap() (map[string]string, error) {
	fmt.Println("Input: Fetching retailer name to retailer code map from metadata ", r.filePath)

	f, err := excelize.OpenFile(r.filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open TSE mapping file: %w", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %w", err)
	}

	dealerNameIdx, err := utils.GetColumnIndex(f, sheetName, "Tally Name(Dealer Name)")
	if err != nil {
		return nil, err
	}
	retailerCodeIdx, err := utils.GetColumnIndex(f, sheetName, "Dealer Code")
	if err != nil {
		return nil, err
	}

	retailerNameToCodeMap := make(map[string]string)
	for _, row := range rows[1:] {
		dealerName := row[dealerNameIdx]
		retailerCode := row[retailerCodeIdx]

		if dealerName == "" {
			continue
		}
		retailerNameToCodeMap[dealerName] = retailerCode
	}

	return retailerNameToCodeMap, nil
}

func (r *ExcelTSEMappingRepository) GetRARetailersMap() (map[string]int, error) {
	fmt.Println("Input: Fetching RA retailers from metadata ", r.filePath)

	f, err := excelize.OpenFile(r.filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open TSE mapping file: %w", err)
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
	typeIdx, err := utils.GetColumnIndex(f, sheetName, "Type")
	if err != nil {
		return nil, err
	}
	countOfRAIdx, err := utils.GetColumnIndex(f, sheetName, "Count of RA")
	if err != nil {
		return nil, err
	}

	raRetailers := make(map[string]int)
	for _, row := range rows[1:] {
		// Check if the row has enough columns
		if len(row) <= countOfRAIdx || len(row) <= typeIdx || len(row) <= dealerCodeIdx {
			continue
		}

		dealerCode := row[dealerCodeIdx]
		retailerType := row[typeIdx]
		countOfRAIdxStr := row[countOfRAIdx]

		// Only include RA retailers
		if retailerType == "RA" && dealerCode != "" {
			countRA, err := strconv.Atoi(countOfRAIdxStr)
			if err != nil {
				fmt.Printf("Warning: Invalid count of RA for retailer %s, skipping...\n", dealerCode)
				continue
			}
			raRetailers[dealerCode] = countRA
		}
	}

	return raRetailers, nil
}
