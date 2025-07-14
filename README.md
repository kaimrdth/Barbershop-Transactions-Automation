# Service Business Payment Processing System

## Overview

This Google Apps Script automates the complex process of transforming raw payment data into detailed financial reports for a service-based business. The system handles commission calculations, tip distribution, product sales tracking, and generates comprehensive reports for both staff compensation and business analytics.

## Problem Solved

Service businesses often struggle with:
- **Manual Commission Calculations**: Tedious and error-prone manual calculation of staff commissions based on varying rates for services vs. products
- **Complex Payment Processing**: Difficulty separating tips, service fees, product sales, and processing costs from mixed payment data
- **Data Consistency**: Maintaining accurate records when dealing with refunds, voids, and duplicate transactions
- **Financial Transparency**: Providing clear breakdowns of earnings for staff and net revenue for the business

## System Architecture

### Data Flow
```
Raw Payment Data → Data Processing → Clean Financial Reports
```

The system operates on four Google Sheets:
1. **Raw**: Unprocessed payment data from external payment processors
2. **Commission Rates**: Staff-specific commission percentages for services and products
3. **Menu of Services**: Service and product pricing lookup table
4. **Processed**: Final output with detailed financial breakdowns

## Key Features

### 🔄 Intelligent Data Processing
- **Duplicate Prevention**: Automatically detects and skips already-processed transactions using Payment ID tracking
- **Error Handling**: Robust validation for missing sheets, malformed data, and edge cases
- **Data Cleaning**: Removes header rows and standardizes data formats

### 💰 Advanced Financial Calculations

#### Commission System
- **Service Commissions**: Calculated based on individual staff rates and service pricing
- **Product Commissions**: Separate rates for retail sales with special handling for business owners
- **Processing Fee Distribution**: Splits payment processing costs between staff and business

#### Smart Parsing
- **Service Recognition**: Extracts service type and staff assignments from combined text fields using regex patterns
- **Product Identification**: Automatically identifies and quantifies product sales from order descriptions
- **Tip Calculation**: Derives tip amounts through financial reconciliation (Amount Paid - Services - Products - Taxes)

### 📊 Comprehensive Reporting

The processed output includes 23 detailed columns:
- **Transaction Details**: Payment ID, date, customer information
- **Service Breakdown**: Service type, staff member, commission rates and amounts
- **Product Analysis**: Product names, quantities, sales amounts, and commissions
- **Financial Summary**: Tips, taxes, discounts, processing fees
- **Business Metrics**: Total staff compensation and net business revenue

### 🎨 Visual Organization
- **Color-coded columns** for different data categories (fees, services, products, adjustments)
- **Automatic formatting** with currency and percentage displays
- **Chronological sorting** for easy data analysis

## Technical Highlights

### Performance Optimization
- **Batch Processing**: Processes multiple transactions in single operations
- **Memory Efficient**: Uses Sets for duplicate detection and Maps for lookup operations
- **Minimal API Calls**: Reduces Google Sheets API usage through range-based operations

### Error Prevention
- **Transaction Deduplication**: Prevents processing the same payment multiple times
- **Status Handling**: Properly zeroes out financial data for refunded/voided transactions
- **Data Validation**: Checks for required fields and handles missing data gracefully

### Business Logic
- **Flexible Commission Structure**: Supports different rates for services vs. products
- **Owner/Manager Rules**: Special handling for business owners with different commission structures
- **Multi-Product Orders**: Intelligently distributes quantities across multiple products in single orders

## Data Processing Logic

### Input Processing
1. **Raw Data Extraction**: Reads payment processor exports with complex multi-column layouts
2. **Field Mapping**: Maps 50+ columns from raw data to relevant business fields
3. **Name Parsing**: Uses regex to extract staff assignments from "Service w/ Staff Name" format

### Financial Calculations
1. **Service Revenue**: `Service Price × Staff Commission Rate = Service Commission`
2. **Product Revenue**: `Product Price × Quantity × Product Commission Rate = Product Commission`
3. **Tip Calculation**: `Amount Paid + Discounts - Services - Products - Tax = Tips`
4. **Net Business**: `Amount Paid - Total Staff Commission - Business Processing Fee - Tax`

### Quality Assurance
- **Duplicate Detection**: Maintains Set of processed Payment IDs
- **Data Integrity**: Validates calculations and handles edge cases
- **Audit Trail**: Comprehensive logging for troubleshooting

## Business Impact

### For Staff
- **Transparent Earnings**: Clear breakdown of service commissions, product commissions, and tips
- **Accurate Payments**: Eliminates manual calculation errors
- **Fair Processing**: Equitable distribution of payment processing costs

### For Business
- **Financial Clarity**: Precise tracking of net revenue after all expenses
- **Operational Efficiency**: Reduces administrative time from hours to minutes
- **Data-Driven Decisions**: Comprehensive reporting enables business analytics

### For Customers
- **Service Quality**: Staff can focus on service rather than payment calculations
- **Transparency**: Clear itemization of services and products

## Technical Stack

- **Platform**: Google Apps Script (JavaScript runtime)
- **Data Storage**: Google Sheets
- **Key Technologies**:
  - Regular Expressions for text parsing
  - Set/Map data structures for performance
  - Batch operations for API efficiency
  - Comprehensive error handling

## Usage

The system runs automatically when new raw data is added, requiring minimal user intervention. The main function `processDataComplete()` handles the entire workflow from data import to formatted output.

---
