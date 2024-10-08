package utils

import (
	"fmt"
	"path/filepath"
	"time"
)

func GenerateOutputPath(outputDir, filePrefix string) string {
	today := time.Now().Format("2006-01-02")
	fileName := fmt.Sprintf("%s_%s/", filePrefix, today)
	return filepath.Join(outputDir, fileName)
}

// ... existing code ...
func GenerateMonthlyOutputPath(outputDir, filePrefix string) string {
	today := time.Now().Format("2006-Jan") // Changed format to YYYY-MM
	fileName := fmt.Sprintf("%s_%s/", filePrefix, today)
	return filepath.Join(outputDir, fileName)
}

// CalculateGrowthPercentage calculates the growth percentage between two values.
func CalculateGrowthPercentage(current, previous float64) int {
	if previous == 0 {
		return 0 // Avoid division by zero
	}
	return int(((current - previous) / previous) * 100) // Cast to int
}
