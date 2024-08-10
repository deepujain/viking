package utils

// PartyData struct to hold the categorized data along with the party name
type PartyData struct {
	RetailerName string
	Amounts      map[string]interface{}
}

// Aggregate bills by retailer and age category
func AggregateCreditByRetailer(bills []Bill, tseMapping map[string]string) map[string]map[string]interface{} {
	ageCategories := []string{"0-7 days", "8-14 days", "15-21 days", "22-30 days", ">30 days"}
	aggregatedCreditRetailer := make(map[string]map[string]interface{})

	// First, group all bills by retailer name
	groupedBills := make(map[string][]Bill)
	for _, bill := range bills {
		groupedBills[bill.RetailerName] = append(groupedBills[bill.RetailerName], bill)
	}

	// Now process each retailer's bills
	for retailerName, partyBills := range groupedBills {
		aggregatedCreditRetailer[retailerName] = make(map[string]interface{})
		totalPendingAmount := 0.0

		// Initialize categories with 0.0
		for _, category := range ageCategories {
			aggregatedCreditRetailer[retailerName][category] = 0.0
		}

		// Categorize bills into age categories and compute total pending amount
		for _, bill := range partyBills {
			totalPendingAmount += bill.PendingAmount
			switch {
			case bill.AgeOfBill >= 0 && bill.AgeOfBill <= 7:
				aggregatedCreditRetailer[retailerName]["0-7 days"] = aggregatedCreditRetailer[retailerName]["0-7 days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill >= 8 && bill.AgeOfBill <= 14:
				aggregatedCreditRetailer[retailerName]["8-14 days"] = aggregatedCreditRetailer[retailerName]["8-14 days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill >= 15 && bill.AgeOfBill <= 21:
				aggregatedCreditRetailer[retailerName]["15-21 days"] = aggregatedCreditRetailer[retailerName]["15-21 days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill >= 22 && bill.AgeOfBill <= 30:
				aggregatedCreditRetailer[retailerName]["22-30 days"] = aggregatedCreditRetailer[retailerName]["22-30 days"].(float64) + bill.PendingAmount
			case bill.AgeOfBill > 30:
				aggregatedCreditRetailer[retailerName][">30 days"] = aggregatedCreditRetailer[retailerName][">30 days"].(float64) + bill.PendingAmount
			}
		}
		aggregatedCreditRetailer[retailerName]["Total"] = totalPendingAmount
		aggregatedCreditRetailer[retailerName]["TSE"] = tseMapping[retailerName] // Add TSE name
	}

	return aggregatedCreditRetailer
}
