package repository

import (
	"fmt"
	"strconv"
	"strings"
	"viking-reports/internal/utils"

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
	fmt.Println("Read the price list given by zonal distributor from ", r.filePath)
	fmt.Println()
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

	var lastType, lastModel, lastColor string                                       // Track last seen values
	var typeIndx, modelIndx, colorIdx, capacityIdx, dlrPriceIdx, mopIdx, mrpIdx int // Track indices
	headerRow := 1
	for i, row := range rows {

		if i == headerRow {
			// Skip header
			// Log the headers for debugging
			fmt.Println("Headers found:", row)

			// Set up column indices from the second row (headers)
			var err error
			typeIndx, err = utils.GetHeaderIndex(f, sheetName, "TYPE", headerRow)
			if err != nil {
				fmt.Println("Error retrieving index for TYPE:", err)
				return nil, err
			}
			modelIndx, err = utils.GetHeaderIndex(f, sheetName, "Model", headerRow)
			if err != nil {
				fmt.Println("Error retrieving index for Model:", err)
				return nil, err
			}
			colorIdx, err = utils.GetHeaderIndex(f, sheetName, "COLOURS", headerRow)
			if err != nil {
				fmt.Println("Error retrieving index for COLOURS:", err)
				return nil, err
			}
			capacityIdx, err = utils.GetHeaderIndex(f, sheetName, "Variant", headerRow) // Added index for capacity
			if err != nil {
				fmt.Println("Error retrieving index for Variant:", err)
				return nil, err
			}
			dlrPriceIdx, err = utils.GetHeaderIndex(f, sheetName, "DLR PRICE", headerRow)
			if err != nil {
				fmt.Println("Error retrieving index for DLR PRICE:", err)
				return nil, err
			}
			mopIdx, err = utils.GetHeaderIndex(f, sheetName, "MOP", headerRow)
			if err != nil {
				fmt.Println("Error retrieving index for MOP:", err)
				return nil, err
			}
			mrpIdx, err = utils.GetHeaderIndex(f, sheetName, "MRP", headerRow)
			if err != nil {
				fmt.Println("Error retrieving index for MRP:", err)
				return nil, err
			}
			// Skip to the next iteration to start processing data rows
			continue
		}
		if len(row) >= 7 { // Ensure we have enough columns
			// Check for merged values
			if row[0] != "" {
				lastType = row[typeIndx]
			}
			if row[1] != "" {
				lastModel = row[modelIndx]
			}
			if row[2] != "" {
				lastColor = row[colorIdx]
			}

			// Use last seen values if current row is missing
			model := lastModel
			color := lastColor
			capacity := row[capacityIdx]
			dlrPrice, _ := strconv.Atoi(row[dlrPriceIdx])
			mop, _ := strconv.Atoi(row[mopIdx])
			mrp, _ := strconv.Atoi(row[mrpIdx])

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

	// Post-process the results to split colors
	return splitColorsInResults(results), nil
	//return results, nil
}

// Helper function to process the results and split colors
func splitColorsInResults(results []PriceListRow) []PriceListRow {
	var finalResults []PriceListRow

	// Known multi-word colors that should be split
	multiWordColors := map[string][]string{
		"SAFARI GREEN MARBLE BLACK":                   {"SAFARI GREEN", "MARBLE BLACK"},
		"GREEN BLACK":                                 {"GREEN", "BLACK"},
		"Sunny Oasis Dark Purple":                     {"Sunny Oasis", "Dark Purple"},
		"TWILIGHT PURPLE WOODLAND GREEN":              {"TWILIGHT PURPLE", "WOODLAND GREEN"},
		"NAVIGATOR BEIGE SUBMARINE BLUE":              {"NAVIGATOR BEIGE", "SUBMARINE BLUE"},
		"NAVIGATOR BEIGE SUBMARINE BLUE EXPLORER RED": {"NAVIGATOR BEIGE", "SUBMARINE BLUE", "EXPLORER RED"},
		"SPEED GREEN DARK PURPLE":                     {"SPEED GREEN", "DARK PURPLE"},
		"VICTORY GOLD SPEED GREEN DARK PURPLE":        {"VICTORY GOLD", "SPEED GREEN", "DARK PURPLE"},
		"MONET GOLD MONET PURPLE EMERALD GREEN":       {"MONET GOLD", "MONET PURPLE", "EMERALD GREEN"},
		"MONET GOLD EMERALD GREEN":                    {"MONET GOLD", "EMERALD GREEN"},
		"FLUID SILVER RAZOR GREEN":                    {"FLUID SILVER", "RAZOR GREEN"},
	}

	// Iterate through each result and split colors
	for _, row := range results {
		// Check if the color is in the known multi-word colors map

		if splitColors, exists := multiWordColors[row.Color]; exists {
			fmt.Println("Found row with no seperator between model colors:", row.Color)
			// If found, split into multiple rows based on the colors provided
			for _, color := range splitColors {
				newRow := row
				newRow.Color = color
				finalResults = append(finalResults, newRow)
			}
		} else {
			// For other colors, check if they need to be split by delimiters
			colors := splitColorsByDelimiters(row.Color)
			for _, color := range colors {
				newRow := row
				newRow.Color = color
				finalResults = append(finalResults, newRow)
			}
		}
	}

	return finalResults
}

// Helper function to split colors based on known delimiters
func splitColorsByDelimiters(color string) []string {
	// Define known separators
	separators := []string{"\n", "/", "\\", ":", ","}

	// Iterate over each separator and split the color string
	for _, sep := range separators {
		if strings.Contains(color, sep) {
			splitColors := strings.Split(color, sep)
			// Trim whitespace from each part
			for i := range splitColors {
				splitColors[i] = strings.TrimSpace(splitColors[i])
			}
			return splitColors
		}
	}

	// If no known separator is found, return the color as-is
	return []string{color}
}
