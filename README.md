# Excel Mapper

A Python library that maps Excel columns to Python object attributes with dynamic access. Read, modify, and save Excel files using clean attribute-based syntax.

## Features

- 📊 **Dynamic Attribute Access**: Access Excel data using clean attribute names
- 🔄 **Automatic Column Sanitization**: Special characters and spaces converted to underscores
- ♻️ **Duplicate Handling**: Automatic suffixing for duplicate column names
- 💾 **Bidirectional Operations**: Read from and save back to Excel files
- 🐍 **Pythonic Interface**: Intuitive object-oriented API
- 📝 **Full Type Hints**: Complete type annotations for better development experience

## Installation

```bash
pip install excel-mapper
```

## Quick Start
```python
from excel_mapper import ExcelMapper

# Load Excel file
mapper = ExcelMapper("financial_data.xlsx")

# Access data using clean attribute names
print(f"First transaction ID: {mapper.rows[0].transaction_id}")
print(f"Transaction amount: {mapper.rows[0].amount_usd}")

# Modify data
mapper.rows[0].transaction_id = "TXN_2024_001"
mapper.rows[0].amount_usd = 1500.75

# Save changes
mapper.save_excel("updated_financial_data.xlsx")
```

## Column Name Conversion

Excel columns are automatically converted to Python-friendly attribute names:


| Excel Column Name | Attribute Name |
|-------------------|----------------|
| `Transaction ID` | `transaction_id` |
| `Amount (USD)` | `amount_usd` |
| `Transaction Date` | `transaction_date` |
| `Category` | `category` |
| `Merchant Name` | `merchant_name` |


**Duplicate columns** get automatic suffixes: `column_a`, `column_b`, etc.

## Advanced Usage

### Access Column Information
```python
# Get column mapping
mapping = mapper.get_column_mapping()
# {`Transaction ID`: `transaction_id`, `Amount (USD)`: `amount_usd`}

# Get original column names
columns = mapper.get_original_columns()
# [`Transaction ID`, `Amount (USD)`]

# Get attribute names
attrs = mapper.get_attribute_names()
# [`transaction_id`, `amount_usd`]
```

### Update Data Programmatically
```python
# Update specific row
mapper.update_row(0, transaction_id="TXN_2024_001", amount_usd=2500.50)

# Update entire column
new_amounts = [amount * 1.1 for amount in mapper.amount_usd]  # 10% increase
mapper.update_column("amount_usd", new_amounts)

# Add new row
mapper.add_row(transaction_id="TXN_2024_999", amount_usd=999.99, category="Office Supplies")
```

### Iterate Through Data
```python
# Iterate through all rows
for row in mapper:
    print(f"Transaction {row.transaction_id}: ${row.amount_usd}")

# Access specific row
row_2 = mapper[1]  # 0-indexed (Excel row 3)
```

## API Reference

### ExcelMapper Class

#### `ExcelMapper(file_path: str)`
Initialize with Excel file path.

#### Methods
- `get_column_mapping() -> Dict[str, str]`: Original → attribute name mapping
- `get_original_columns() -> List[str]`: Original column names
- `get_attribute_names() -> List[str]`: Sanitized attribute names
- `save_excel(file_path=None, overwrite=False)`: Save to Excel file
- `update_row(row_index: int, **kwargs)`: Update specific row
- `update_column(column_attr: str, new_values: List)`: Update entire column
- `add_row(**kwargs)`: Add new row

### Row Objects
Each row provides dynamic attribute access:
```python
row.transaction_id        # Access data
row.amount_usd = 100.50   # Modify data
row.to_dict()             # Convert to dictionary
```

## Examples

### Basic Data Processing
```python
mapper = ExcelMapper("expense_report.xlsx")

total_expenses = 0
for row in mapper:
    if row.category == "Travel":
        total_expenses += row.amount_usd
        row.approval_status = "Pending Manager Review"

print(f"Total travel expenses: ${total_expenses}")
mapper.save_excel("processed_expenses.xlsx")
```

### Data Analysis
```python
mapper = ExcelMapper("quarterly_sales.xlsx")

q1_sales = sum(row.sales_amount for row in mapper if row.quarter == "Q1")
top_performers = [row for row in mapper if row.sales_amount > 100000]
```

## Requirements

- Python 3.7+
- pandas >= 1.0
- openpyxl >= 3.0

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/new-feature`
3. Commit changes: `git commit -am 'Add new feature'`
4. Push to branch: `git push origin feature/new-feature`
5. Submit a pull request

## License

MIT License - see LICENSE file for details.

## Support

- Documentation: https://github.com/spllat-00/excel-mapper/wiki
- Issue Tracker: https://github.com/spllat-00/excel-mapper/issues
- Discussions: https://github.com/spllat-00/excel-mapper/discussions

## Versioning

This project uses Semantic Versioning. Current version: 1.0.0

---

**Excel Mapper** - Making Excel data manipulation in Python more intuitive and Pythonic!
