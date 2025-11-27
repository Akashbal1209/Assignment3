## NSE Options Trading Analyzer
A Python-based tool for analyzing NSE (National Stock Exchange) options trading opportunities. This analyzer calculates optimal strike prices, premiums, and Internal Rate of Return (IRR) for both Call and Put options across multiple stocks and indices.
Features

Automated Strike Price Calculation: Finds optimal CE (Call) and PE (Put) strike prices based on configurable margin percentages
IRR Analysis: Calculates annualized returns for option writing strategies
52-Week Percentile Tracking: Shows current price position relative to 52-week range
Premium Estimation: Generates realistic option premium estimates
Excel Reports: Exports detailed analysis to formatted Excel files
Visual Analytics: Creates comprehensive dashboard with 6 different charts
Configurable Parameters: Customize margin requirements and lot sizes

## Installation
Prerequisites
bashpython 3.7+
Required Libraries
bashpip install pandas matplotlib openpyxl
Or install all dependencies at once:
bashpip install -r requirements.txt

## requirements.txt
pandas>=1.3.0
matplotlib>=3.4.0
openpyxl>=3.0.9
Usage
Basic Usage
pythonfrom options_analyzer import OptionsAnalyzer

# Create analyzer instance
analyzer = OptionsAnalyzer()

# Run analysis with default parameters (15% margin, 1x lot size)
analyzer.run()
Custom Parameters
python# Run with custom margin and lot multiplier
analyzer.run(margin_percent=20, lot_multiplier=2)
Command Line Execution
bashpython options_analyzer.py

## Configuration Parameters

| Parameter | Description | Default | Example |
|-----------|-------------|---------|---------|
| `margin_percent` | Safety margin from current strike price | 15% | 10, 15, 20 |
| `lot_multiplier` | Multiplier for standard lot sizes | 1 | 1, 2, 5 |

## Output Files

The analyzer generates two files with timestamps:

1. **Excel Report**: `Options_Analysis_YYYYMMDD_HHMMSS.xlsx`
   - Complete analysis data in tabular format
   - Auto-formatted columns
   - Ready for further analysis

2. **Visual Dashboard**: `Options_Analysis_Graphs_YYYYMMDD_HHMMSS.png`
   - 6 comprehensive charts showing different aspects of the analysis

## Analysis Components

### 1. Strike Price Calculation
- Finds nearest standard strike (₹50 intervals)
- Applies configurable margin for safety
- CE Strike: Below current price by margin %
- PE Strike: Above current price by margin %

### 2. Premium Estimation
- ITM (In-The-Money): Intrinsic value + time value
- OTM (Out-of-The-Money): Time value based on spot price

### 3. IRR Calculation
```
IRR = (Premium / Margin Required) × (365 / Days to Expiry) × 100
4. Percentile Analysis
Shows where current price stands in the 52-week range:

>50%: Upper half of range (potential resistance)
<50%: Lower half of range (potential support)

Dashboard Visualizations

Call vs Put IRR Comparison: Side-by-side IRR comparison for all symbols
Premium Comparison: CE and PE premium analysis
52-Week Price Percentile: Visual indication of price positioning
Spot Price vs Strikes: Shows relationship between current price and calculated strikes
Average IRR Comparison: Overall performance of Call vs Put strategies
Premium to Strike Ratio: Efficiency metric for option selection

Sample Output Structure
Excel Columns

Symbol
Spot Price
52W High/Low
Percentile
Lot Size
CE Strike & Premium & IRR
PE Strike & Premium & IRR
Margin Used (%)

Included Symbols
The analyzer currently includes:
Indices:

NIFTY
BANKNIFTY

Stocks:

RELIANCE
TCS
INFY
HDFCBANK
ICICIBANK
SBIN
BHARTIARTL
HINDUNILVR

Customization
Adding New Symbols
Modify the get_sample_data() method:
pythondef get_sample_data(self):
    return [
        {
            'symbol': 'SYMBOL_NAME',
            'spot_price': 1000,
            'high_52w': 1200,
            'low_52w': 800,
            'lot_size': 500
        },
        # Add more symbols...
    ]
Integration with Live Data
Replace get_sample_data() with API calls to:

NSE India API
Yahoo Finance
Other market data providers

pythondef get_sample_data(self):
    # Replace with actual API call
    # response = requests.get('YOUR_API_ENDPOINT')
    # return response.json()
    pass

## Use Cases

Option Writing Strategies: Identify high IRR opportunities
Risk Assessment: Evaluate safety margins before selling options
Portfolio Analysis: Compare multiple symbols simultaneously
Market Positioning: Understand price levels relative to 52-week range


Margin requirement: 15% of strike price (configurable)
Days to expiry: 30 days (for IRR calculation)
Strike intervals: ₹50
Premium calculations are simplified models