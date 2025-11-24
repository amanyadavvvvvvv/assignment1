# NSE Stock Analyzer

A Python script that analyzes NSE stocks and calculates their current position relative to 52-week, 3-month, and 1-month price ranges.

## Features

- Fetches real-time stock data from Yahoo Finance
- Calculates 52-week, 3-month, and 1-month high/low prices
- Determines current price position as percentage
- Generates formatted Excel report with analysis

## Requirements

```bash
pip install -r requirements.txt
```

**Dependencies:**
- pandas
- openpyxl
- yfinance

## Usage

```bash
python main.py
```

## Output

The script generates:
- Console output with stock analysis summary
- Excel file: `Stock_Analysis_Report_[timestamp].xlsx`

## Stocks Analyzed

- IDEA (Vodafone Idea Limited)
- ADANIPORTS (Adani Ports and SEZ)
- RELIANCE (Reliance Industries)
- BAJAJ-AUTO (Bajaj Auto Limited)

## Customization

Edit the `stock_dict` in the `main()` function to analyze different stocks:

```python
stock_dict = {
    'SYMBOL': 'Company Name',
    'TCS': 'Tata Consultancy Services'
}
```

## Notes

- Requires active internet connection
- Data source: Yahoo Finance
- 2-second delay between stock fetches to avoid rate limiting