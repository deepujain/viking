package report

import (
	"fmt"
	"sort"
	"viking-reports/internal/config"
	"viking-reports/internal/repository"
	"viking-reports/internal/utils"
	"viking-reports/pkg/excel"
)

type GrowthReportGenerator struct {
	cfg            *config.Config
	salesRepo      repository.SalesRepository
	tseMappingRepo repository.TSEMappingRepository
}

func NewGrowthReportGenerator(cfg *config.Config) *GrowthReportGenerator {
	return &GrowthReportGenerator{
		cfg:            cfg,
		salesRepo:      repository.NewExcelSalesRepository(cfg.ReportFiles.GrowthReport.MTDSO), // Use one of the file paths
		tseMappingRepo: repository.NewExcelTSEMappingRepository(cfg.CommonFiles.TSEMapping),
	}
}

func (g *GrowthReportGenerator) Generate() error {
	fmt.Println("Generating Growth report...")

	mtdSOData, err := g.salesRepo.GetSellData("mtdSO")
	if err != nil {
		return fmt.Errorf("error reading MTD SO data: %w", err)
	}

	lmtdSOData, err := g.salesRepo.GetSellData("lmtdSO")
	if err != nil {
		return fmt.Errorf("error reading LMTD SO data: %w", err)
	}

	mtdSTData, err := g.salesRepo.GetSellData("mtdST")
	if err != nil {
		return fmt.Errorf("error reading MTD ST data: %w", err)
	}

	lmtdSTData, err := g.salesRepo.GetSellData("lmtdST")
	if err != nil {
		return fmt.Errorf("error reading LMTD ST data: %w", err)
	}

	tseMapping, err := g.tseMappingRepo.GetRetailerCodeToTSEMap()
	if err != nil {
		return fmt.Errorf("error reading TSE mapping: %w", err)
	}

	report := g.generateGrowthReport(mtdSOData, lmtdSOData, mtdSTData, lmtdSTData)

	outputPath := utils.GenerateOutputPath(g.cfg.OutputDir, "daily_growth_report")
	if err := g.writeGrowthReport(outputPath, report, tseMapping); err != nil {
		return fmt.Errorf("error writing growth report: %w", err)
	}

	fmt.Printf("Growth report generated successfully: %s\n", outputPath)
	return nil
}

func (g *GrowthReportGenerator) generateGrowthReport(mtdSOData, lmtdSOData, mtdSTData, lmtdSTData map[string]*repository.SellData) []repository.GrowthReport {
	var report []repository.GrowthReport

	for dealerCode := range mtdSOData {
		mtdSO := g.getOrCreateSellData(mtdSOData, dealerCode)
		lmtdSO := g.getOrCreateSellData(lmtdSOData, dealerCode)
		mtdST := g.getOrCreateSellData(mtdSTData, dealerCode)
		lmtdST := g.getOrCreateSellData(lmtdSTData, dealerCode)

		reportEntry := repository.GrowthReport{
			DealerCode:  dealerCode,
			DealerName:  mtdSO.DealerName,
			MTDSO:       mtdSO.MTDS,
			LMTDSO:      lmtdSO.MTDS,
			GrowthSOPct: utils.CalculateGrowthPercentage(float64(mtdSO.MTDS), float64(lmtdSO.MTDS)), // Convert to float64
			MTDST:       mtdST.MTDS,
			LMTDST:      lmtdST.MTDS,
			GrowthSTPct: utils.CalculateGrowthPercentage(float64(mtdST.MTDS), float64(lmtdST.MTDS)), // Convert to float64
		}
		report = append(report, reportEntry)
	}

	sort.Slice(report, func(i, j int) bool {
		return report[i].GrowthSOPct > report[j].GrowthSOPct
	})

	return report
}

func (g *GrowthReportGenerator) getOrCreateSellData(data map[string]*repository.SellData, dealerCode string) *repository.SellData {
	if data, exists := data[dealerCode]; exists {
		return data
	}
	return &repository.SellData{DealerCode: dealerCode, DealerName: "", MTDS: 0}
}

func (g *GrowthReportGenerator) writeGrowthReport(outputPath string, report []repository.GrowthReport, tseMapping map[string]string) error {
	f := excel.NewFile()
	sheetName := "Growth Report"

	headers := []string{"Dealer Code", "Dealer Name", "TSE", "MTD SO", "LMTD SO", "Growth SO%", "MTD ST", "LMTD ST", "Growth ST%"}
	if err := excel.WriteHeaders(f, sheetName, headers); err != nil {
		return err
	}

	row := 2
	for _, entry := range report {
		cellData := []interface{}{
			entry.DealerCode,
			entry.DealerName,
			tseMapping[entry.DealerCode],
			entry.MTDSO,
			entry.LMTDSO,
			entry.GrowthSOPct,
			entry.MTDST,
			entry.LMTDST,
			entry.GrowthSTPct,
		}
		if err := excel.WriteRow(f, sheetName, row, cellData); err != nil {
			return err
		}
		row++
	}

	return f.SaveAs(outputPath)
}
