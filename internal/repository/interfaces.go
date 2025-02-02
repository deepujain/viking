package repository

type TSEMappingRepository interface {
	GetRARetailersMap() (map[string]int, error)
	GetRetailerCodeToTSEMap() (map[string]string, error)
	GetRetailerCodeToNameMap() (map[string]string, error)
	GetRetailerNameToTSEMap(dealerNameHeader string) (map[string]string, error)
	GetRetailerNameToCodeMap() (map[string]string, error)
}

type ProductPriceRepository interface {
	GetProductPrices() (map[string]float64, error)
}

type InventoryRepository interface {
	ComputeInventoryShortFall() (map[string]*InventoryShortFallRepo, error)
	ComputeMaterialModelCount() (map[string]*ModelCountRepo, error)
	ComputeDealerSPUInventory(modelsOfInterest map[string]struct{}) (map[string]*SPUInventoryCount, error)
	ComputeRADealerSPUInventory(modelsOfInterest map[string]struct{}, raRetailers map[string]int) (map[string]*SPUInventoryCount, error)
}

type CreditRepository interface {
	GetCreditData() (map[string]*CreditData, error)
	GetBills() ([]Bill, error)
	AggregateCreditByRetailer(bills []Bill, tseMapping map[string]string, retailerNameToCodeMap map[string]string) map[string]map[string]interface{}
}

type DebitRepository interface {
	GetDebit() (map[string]float64, error)
}

type SalesRepository interface {
	GetSales(fileType string) (map[string]*SellData, error)
	GetDealerSPUSales(salesFilePath string, modelsOfInterest map[string]struct{}) (map[string]*DealerSPUSales, error)
}

type PriceListRepository interface {
	GetPriceListData() ([]PriceListRow, error)
	GetMaterialCodeMap() (map[string]int, error)
}
type SalesTargetRepository interface {
	ReadSales(fileType string, tseMap map[string]string) ([]*SalesData, error)
}
