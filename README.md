# Business Transaction Processing System

Here's what happens in simple terms:

1. **Raw transaction data comes in** - timestamps, payment IDs, amounts, customer names, but the details are all jumbled together in confusing formats
2. **The system reads and cleans this data** - it figures out which staff member did what service, what products were sold, and calculates all the financial details
3. **It produces a clean, organized report** that shows:
   - How much each staff member earned in commissions
   - What services and products were sold
   - How much the business kept after paying commissions and fees
   - Tips, taxes, discounts, and other financial details
   - Customer information for relationship tracking

---

## System Overview (Technical)

This Google Apps Script-based system consists of two main components that work together to process business transaction data:

### Core Components

1. **Data Import & Preprocessing** (`processTransactionData()`)
2. **Business Logic Processing** (`processBusinessTransactions()`)

### Data Flow Architecture

```
Raw Payment Data → Data Cleaning → Business Processing → Formatted Reports
```

The system integrates data from multiple sources:
- Transaction records from payment processors
- Commission rate tables
- Service/product pricing menus
- Customer information

---

## Detailed Technical Documentation

### Component 1: Data Import & Preprocessing

**File:** `data-import-processor.js`
**Function:** `processTransactionData()`

#### Purpose
Handles the initial ingestion and cleaning of raw transaction data from payment processors.

#### Key Features
- **Duplicate Prevention:** Maintains a registry of processed transaction IDs to prevent double-processing
- **Data Validation:** Checks for minimum required data before processing
- **Header Management:** Automatically sets up or validates expected column headers
- **Data Cleaning:** Removes header rows and formats data consistently

#### Process Flow
1. Retrieves raw transaction data from the "Raw Data" sheet
2. Removes duplicate header rows (assumes 2-row header structure)
3. Checks against existing processed transaction IDs
4. Maps raw columns to standardized format:
   ```javascript
   // Column mapping example
   date: row[0],           // Timestamp
   transactionID: row[1],  // Unique identifier
   currency: row[3],       // Currency code
   amount: row[4],         // Transaction amount
   // ... additional mappings
   ```
5. Calculates basic metrics (commission, fees, net amounts)
6. Appends new data to the "Processed Data" sheet
7. Sorts by timestamp (most recent first)

#### Error Handling
- Validates minimum data requirements
- Handles missing or malformed data gracefully
- Logs processing statistics

---

### Component 2: Business Logic Processing

**File:** `business-transaction-processor.js`
**Function:** `processBusinessTransactions()`

#### Purpose
Transforms cleaned transaction data into detailed business intelligence with commission calculations, product sales tracking, and financial reporting.

#### Advanced Features

##### Multi-Source Data Integration
```javascript
// Integrates data from multiple sheets
const commissionSheet = ss.getSheetByName('Commission Rates');
const menuSheet = ss.getSheetByName('Service Menu');
```

##### Intelligent Service Description Parsing
The system includes sophisticated parsing logic to extract structured data from unstructured service descriptions:

```javascript
function parseServiceDescription(serviceDescription, quantityRaw) {
  // Parses formats like: "Haircut w/ John, Hair Product, Additional Fees"
  // Extracts: staff name, service type, products, and fees
}
```

**Parsing Capabilities:**
- Service + Staff patterns: `"Service Type w/ Staff Name"`
- Product identification and quantity distribution
- Additional fee detection
- Automatic fallback for product-only transactions

##### Commission Calculation Engine
```javascript
// Service commission
let staffServiceCommission = servicePrice * serviceCommissionRate;

// Product commission (with owner exceptions)
if (businessOwners.includes(staffName)) {
  productCommissionRate = 0; 
} else {
  productCommissionRate = commissionRates.productRate;
}
```

**Commission Features:**
- Separate rates for services vs. products
- Special handling for business owners
- Processing fee splitting between staff and business
- Tip calculation and allocation

##### Financial Calculations
The system performs complex financial modeling:

```javascript
// Tip calculation algorithm
let tips = amountPaid + discounts - servicePrice - productSales - productTax;

// Net business revenue calculation
let netBusinessRevenue = amountPaid - totalStaffCommission - businessProcessingFee - productTax;
```

##### Product Quantity Distribution
When multiple products are sold in a single transaction:
```javascript
// Intelligent quantity distribution among products
if (products.length === 1) {
  products[0].quantity = actualProductQuantity;
} else {
  const perProductQuantity = Math.floor(actualProductQuantity / products.length);
  // Distributes remainder to first products
}
```

#### Data Quality Management

##### Duplicate Detection & Removal
```javascript
function removeDuplicateTransactionIDs() {
  // Removes duplicates while preserving first occurrence
  // Operates on Transaction ID as primary key
}
```

##### Status-Based Processing
- Handles refunds and voids by zeroing out financial calculations
- Maintains transaction history for audit purposes
- Preserves original status information

#### Automated Formatting System

The system includes comprehensive formatting automation:

```javascript
function applyBusinessFormatting(sheet, headerLength) {
  // Color-coded column groups:
  // - Processing fees: Light blue
  // - Service data: Light green  
  // - Product data: Light yellow
  // - Adjustments: Light pink
  
  // Automatic number formatting:
  // - Currency columns: $#,##0.00
  // - Percentage columns: 0.00%
}
```

---

## Data Schema

### Input Data (Raw Transactions)
| Column | Description | Type |
|--------|-------------|------|
| A | Date & Time | DateTime |
| B | Transaction ID | String |
| D | Amount | Currency |
| F | Processing Fee | Currency |
| J | Status | String |
| K | Customer Name | String |
| M | Service Description | String |
| N | Quantity | Integer |

### Output Data (Processed Transactions)
| Column | Field | Description |
|--------|-------|-------------|
| A | Transaction ID | Unique identifier |
| B | Date & Time | Transaction timestamp |
| C | Service Type | Parsed service category |
| D | Staff Member | Assigned staff member |
| E | Additional Fees | Fee indicator |
| F-U | Financial Data | Comprehensive financial breakdown |
| V | Status | Transaction status |
| W | Customer Name | Customer information |

---

## Configuration

### Commission Rates Sheet
Set up staff commission rates:
```
Staff Name | Service Rate | Product Rate
John Smith | 0.45        | 0.10
Jane Doe   | 0.50        | 0.15
```

### Service Menu Sheet
Define service and product pricing:
```
Item Name          | Price
Haircut           | 45.00
Hair Product      | 25.00
Additional Fees   | 10.00
```

---

## Usage

### Initial Setup
1. Create required sheets: "Raw Transactions", "Processed Transactions", "Commission Rates", "Service Menu"
2. Configure commission rates and menu pricing
3. Import raw transaction data

### Running the System
1. **Data Import:** Run `processTransactionData()` to clean and import new transactions
2. **Business Processing:** Run `processBusinessTransactions()` to generate business intelligence
3. **Review Reports:** Analyze the formatted output in "Processed Transactions"

### Automation
Both functions can be scheduled to run automatically using Google Apps Script triggers.

---

## Key Benefits

### For Business Owners
- **Automated Commission Calculations:** No more manual spreadsheet work
- **Real-time Business Metrics:** Instant visibility into revenue, costs, and profitability
- **Staff Performance Tracking:** Detailed commission and sales data per team member
- **Customer Relationship Data:** Organized customer transaction history

### For Technical Teams
- **Scalable Architecture:** Handles growing transaction volumes
- **Maintainable Code:** Well-documented, modular design
- **Error Resilience:** Comprehensive error handling and data validation
- **Flexible Configuration:** Easy to adjust commission rates and business rules

---

## Technical Requirements

- Google Workspace account with Sheets and Apps Script access
- Basic understanding of spreadsheet structure for configuration
- Access to raw transaction data from payment processor

---

## Future Enhancements

- **API Integration:** Direct connection to payment processors
- **Advanced Analytics:** Trend analysis and forecasting
- **Mobile Dashboard:** Real-time mobile reporting interface
- **Multi-location Support:** Handling multiple business locations
