package excel

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// NewFile creates a new Excel file
func NewFile() *excelize.File {
	return excelize.NewFile()
}

// WriteHeaders writes the headers to the Excel sheet
func WriteHeaders(f *excelize.File, sheetName string, headers []string) error {
	headerStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"FFFF00"}, Pattern: 1},
		Font: &excelize.Font{Bold: true},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create header style: %w", err)
	}

	for i, header := range headers {
		cell := fmt.Sprintf("%s%d", string('A'+i), 1)
		f.SetCellValue(sheetName, cell, header)
		f.SetCellStyle(sheetName, cell, cell, headerStyle)
	}

	return nil
}

// WriteRow writes a row of data to the Excel sheet
func WriteRow(f *excelize.File, sheetName string, rowIndex int, data []interface{}) error {
	borderStyle, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})

	for i, value := range data {
		cell := fmt.Sprintf("%c%d", 'A'+i, rowIndex)
		f.SetCellValue(sheetName, cell, value)
		if err := f.SetCellStyle(sheetName, cell, cell, borderStyle); err != nil {
			return fmt.Errorf("error setting style for cell %s: %w", cell, err)
		}
	}
	return nil
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

// AdjustColumnWidths adjusts the width of columns in the Excel sheet
func AdjustColumnWidths(f *excelize.File, sheetName string) {
	cols, _ := f.GetCols(sheetName)
	for i, col := range cols {
		maxWidth := 0
		for _, cell := range col {
			if len(cell) > maxWidth {
				maxWidth = len(cell)
			}
		}
		// Reduce the width by 10%
		finalWidth := float64(maxWidth) * 0.9

		f.SetColWidth(sheetName, string(rune('A'+i)), string(rune('A'+i)), finalWidth+2)
	}
}
