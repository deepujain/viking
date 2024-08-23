package utils

import (
	"strconv"
	"strings"
)

// ParseFloat parses a string to float64, returning 0 if parsing fails
func ParseFloat(s string) float64 {
	f, err := strconv.ParseFloat(strings.TrimSpace(s), 64)
	if err != nil {
		return 0
	}
	return f
}

// ParseInt parses a string to int, returning 0 if parsing fails
func ParseInt(s string) int {
	i, err := strconv.Atoi(strings.TrimSpace(s))
	if err != nil {
		return 0
	}
	return i
}
