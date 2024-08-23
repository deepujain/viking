package utils

import (
	"fmt"
	"os"
	"time"
)

// GetTodayFormatted returns today's date in the format "2006-01-02"
func GetTodayFormatted() string {
	return time.Now().Format("2006-01-02")
}

// CreateDateFolder creates a folder with today's date
func CreateDateFolder(baseDir string) (string, error) {
	today := GetTodayFormatted()
	dirPath := fmt.Sprintf("%s_%s", baseDir, today)
	if err := os.MkdirAll(dirPath, os.ModePerm); err != nil {
		return "", fmt.Errorf("failed to create directory: %w", err)
	}
	return dirPath, nil
}
