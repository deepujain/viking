package utils

import (
	"strconv"
	"strings"
	"time"
)

// Utility function to trim and clean strings
func CleanString(s string) string {
	return strings.TrimSpace(s)
}

// Utility function to safely convert a string to a float64
func SafeParseFloat(s string) (float64, error) {
	s = CleanString(s)
	value, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0.0, err
	}
	return value, nil
}

// Utility function to safely convert a string to an int
func SafeParseInt(s string) (int, error) {
	s = CleanString(s)
	value, err := strconv.Atoi(s)
	if err != nil {
		return 0, err
	}
	return value, nil
}

// Function to get today’s date formatted as "2006-01-02"
func GetTodayDate() string {
	return time.Now().Format("2006-01-02")
}
