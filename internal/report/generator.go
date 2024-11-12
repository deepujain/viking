package report

import (
	"fmt"
	"viking-reports/internal/config"
)

type ReportGenerator interface {
	Generate() error
}

func NewReportGenerator(reportType string, cfg *config.Config) (ReportGenerator, error) {
	switch reportType {
	case "cogs":
		return NewCOGSReportGenerator(cfg), nil
	case "credit":
		return NewCreditReportGenerator(cfg), nil
	case "growth":
		return NewGrowthReportGenerator(cfg), nil
	case "pricelist":
		return NewPriceListGenerator(cfg), nil
	case "salestarget":
		return NewSalesTargetGenerator(cfg), nil
	case "zso":
		return NewZSOReportGenerator(cfg), nil
	case "ranorms":
		return NewRANormsReportGenerator(cfg), nil
	default:
		return nil, fmt.Errorf("unknown report type: %s", reportType)
	}
}
