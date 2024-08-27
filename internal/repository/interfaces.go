package repository

type TSEMappingRepository interface {
	GetRetailerCodeToTSEMap() (map[string]string, error)
	GetRetailerNameToTSEMap() (map[string]string, error)
	GetRetailerNameToCodeMap() (map[string]string, error)
}

type ProductPriceRepository interface {
	GetProductPrices() (map[string]float64, error)
}

type InventoryRepository interface {
	ComputeInventoryShortFall() (map[string]*InventoryShortFallRepo, error)
	ComputeMaterialModelCount() (map[string]*ModelCountRepo, error)
}

type CreditRepository interface {
	GetCreditData() (map[string]*CreditData, error)
	GetBills() ([]Bill, error)
	AggregateCreditByRetailer(bills []Bill, tseMapping map[string]string, retailerNameToCodeMap map[string]string) map[string]map[string]interface{}
}

type SalesRepository interface {
	GetSellData(fileType string) (map[string]*SellData, error)
}
