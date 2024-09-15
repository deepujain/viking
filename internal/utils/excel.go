package utils

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// ReadExcelFile reads an Excel file and returns the file object
func ReadExcelFile(filePath string) (*excelize.File, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file %s: %w", filePath, err)
	}
	return f, nil
}

// GetColumnIndex returns the index of a column given its name
func GetColumnIndex(f *excelize.File, sheetName, columnName string) (int, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return -1, fmt.Errorf("failed to read rows: %w", err)
	}
	if len(rows) == 0 {
		return -1, fmt.Errorf("no rows found in sheet")
	}
	for i, col := range rows[0] {
		if col == columnName {
			return i, nil
		}
	}
	return -1, fmt.Errorf("column %s not found", columnName)
}

// GetHeaderIndex returns the index of a column given its name
func GetHeaderIndex(f *excelize.File, sheetName, columnName string, headerRow int) (int, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return -1, fmt.Errorf("failed to read rows: %w", err)
	}
	if len(rows) == 0 {
		return -1, fmt.Errorf("no rows found in sheet")
	}
	for i, col := range rows[headerRow] {
		if col == columnName {
			return i, nil
		}
	}
	return -1, fmt.Errorf("column %s not found", columnName)
}

// WriteExcelFile saves the Excel file to the specified path
func WriteExcelFile(f *excelize.File, filePath string) error {
	if err := f.SaveAs(filePath); err != nil {
		return fmt.Errorf("failed to save file: %w", err)
	}
	return nil
}
