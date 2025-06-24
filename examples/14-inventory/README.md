# Inventory Management System Example

This example demonstrates how to create comprehensive inventory management reports using the Go Excel Builder library.

## Features Demonstrated

### Inventory Tracking
- Product catalog management
- Stock level monitoring
- Supplier information
- Location tracking
- Serial number management

### Reporting
- Current inventory status
- Low stock alerts
- Inventory valuation
- Movement history
- Reorder recommendations

### Analytics
- ABC analysis
- Turnover rates
- Seasonal trends
- Cost analysis
- Performance metrics

### Advanced Features
- Barcode integration
- Multi-location support
- Batch tracking
- Expiration date monitoring
- Automated calculations

## Generated Files

Running this example will create several Excel files:

1. `14-inventory-current.xlsx` - Current inventory status
2. `14-inventory-movements.xlsx` - Stock movement history
3. `14-inventory-analysis.xlsx` - Inventory analysis and metrics
4. `14-inventory-reports.xlsx` - Management reports
5. `14-inventory-alerts.xlsx` - Alerts and notifications

## Usage

```bash
go run main.go
```

## Expected Output

The program will generate comprehensive inventory management reports including:

- **Current Stock**: Real-time inventory levels with location details
- **Movement History**: Complete audit trail of all stock movements
- **ABC Analysis**: Classification of items by value and importance
- **Reorder Reports**: Automated reorder point calculations
- **Valuation Reports**: Current inventory value by various methods
- **Performance Metrics**: Turnover rates, carrying costs, and efficiency metrics
- **Alert Dashboard**: Low stock, expiration, and other critical alerts

## Key Learning Points

- Inventory data modeling and organization
- Multi-sheet workbook management
- Conditional formatting for alerts
- Formula-based calculations
- Data validation and integrity
- Professional report formatting
- Dashboard creation techniques
- Automated alert systems

## Configuration

The example includes configurable parameters for:
- Reorder points and safety stock levels
- ABC analysis thresholds
- Alert criteria and thresholds
- Valuation methods (FIFO, LIFO, Average)
- Reporting periods and frequencies

## Integration Notes

This example can be easily integrated with:
- ERP systems
- Barcode scanners
- POS systems
- Supplier databases
- Warehouse management systems