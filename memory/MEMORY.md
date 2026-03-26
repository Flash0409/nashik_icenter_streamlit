# Nashik iCenter Forecast & Stock Analysis Tool

## Data Structure Understanding
- **Orderbook File** (OpenOrdersBOM sheet): BOM data with Project Name, Work Order Number, Component Code, Required Quantity
- **Forecast File** (FY26 sheet): Planning data with Project Name, Cabinets Qty, Build/Ship periods
- **Stock File**: Two sheets - "Stock" (inventory) and "Open Order" (incoming POs)

## Key Issue
Multiple work orders for same project don't consolidate properly:
- One project may have 5-500 work orders
- Components should be grouped by ITEM CODE across all work orders
- Need to sum required quantities and compare against available stock

## Grouping Strategy
1. Group by ITEM CODE (primary aggregation key)
2. Show breakdown by Work Order Number
3. Aggregate all WO quantities for same component
4. Compare total against Stock + Open PO Items

## App Structure (Updated)
- Enhanced aggregation in Step 4
- Add Work Order breakdown details
- Better visualization of shortages
- Export detailed grouping report
