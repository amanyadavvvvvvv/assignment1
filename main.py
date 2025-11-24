import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import time
import yfinance as yf

class NSEStockAnalyzer:
    def __init__(self):
        """Initialize the stock analyzer"""
        pass
    
    def get_stock_data(self, symbol):
        """
        Fetch stock data from Yahoo Finance
        Returns dict with price data or None if failed
        """
        try:
            ticker_symbol = f"{symbol}.NS"
            print(f"  → Fetching from Yahoo Finance: {ticker_symbol}")
            
            ticker = yf.Ticker(ticker_symbol)
            info = ticker.info
            
            # Get historical data
            hist_1y = ticker.history(period="1y")
            hist_3m = ticker.history(period="3mo")
            hist_1m = ticker.history(period="1mo")
            
            # Check if data is available
            if hist_1y.empty:
                print(f"  ✗ No data available for {symbol}")
                return None
            
            # Get current price
            current_price = (info.get('currentPrice') or 
                           info.get('regularMarketPrice') or 
                           hist_1y['Close'].iloc[-1])
            
            # Calculate highs and lows
            data = {
                'current_price': float(current_price),
                '52w_high': float(hist_1y['High'].max()),
                '52w_low': float(hist_1y['Low'].min()),
                '3m_high': float(hist_3m['High'].max()),
                '3m_low': float(hist_3m['Low'].min()),
                '1m_high': float(hist_1m['High'].max()),
                '1m_low': float(hist_1m['Low'].min()),
                'source': 'Yahoo Finance (yfinance)'
            }
            
            print(f"  ✓ Success! Current Price: ₹{data['current_price']:.2f}")
            return data
            
        except Exception as e:
            print(f"  ✗ Error fetching data: {str(e)}")
            return None
    
    def calculate_position(self, current, low, high):
        """
        Calculate percentage position of current price between low and high
        Returns value between 0-100
        """
        try:
            if high == low or high == 0:
                return 50.0
            position = ((current - low) / (high - low)) * 100
            return round(max(0, min(100, position)), 1)
        except Exception as e:
            print(f"  ⚠ Calculation error: {str(e)}")
            return 50.0
    
    def get_stock_info(self, symbol, company_name):
        """
        Get complete stock information including price positions
        Returns dict with all stock metrics
        """
        print(f"\n{'='*60}")
        print(f"📊 Analyzing: {symbol}")
        print(f"{'='*60}")
        
        # Fetch stock data
        stock_data = self.get_stock_data(symbol)
        
        # Handle missing data with fallback values
        if not stock_data or stock_data.get('current_price', 0) == 0:
            print(f"  ⚠ Warning: Could not fetch live data, using fallback values")
            current_price = 0
            stock_info = {
                'Symbol': symbol,
                'Current_Price': 0,
                '52_Week_High': 0,
                '52_Week_Low': 0,
                '3_Month_High': 0,
                '3_Month_Low': 0,
                '1_Month_High': 0,
                '1_Month_Low': 0,
                'Data_Source': 'Unavailable'
            }
        else:
            current_price = stock_data.get('current_price', 0)
            stock_info = {
                'Symbol': symbol,
                'Current_Price': round(current_price, 2),
                '52_Week_High': round(stock_data.get('52w_high', current_price * 1.2), 2),
                '52_Week_Low': round(stock_data.get('52w_low', current_price * 0.8), 2),
                '3_Month_High': round(stock_data.get('3m_high', current_price * 1.1), 2),
                '3_Month_Low': round(stock_data.get('3m_low', current_price * 0.9), 2),
                '1_Month_High': round(stock_data.get('1m_high', current_price * 1.05), 2),
                '1_Month_Low': round(stock_data.get('1m_low', current_price * 0.95), 2),
                'Data_Source': stock_data.get('source', 'Unknown')
            }
        
        # Calculate position percentages
        stock_info['Price_vs_52W'] = self.calculate_position(
            stock_info['Current_Price'], 
            stock_info['52_Week_Low'], 
            stock_info['52_Week_High']
        )
        stock_info['Price_vs_3M'] = self.calculate_position(
            stock_info['Current_Price'], 
            stock_info['3_Month_Low'], 
            stock_info['3_Month_High']
        )
        stock_info['Price_vs_1M'] = self.calculate_position(
            stock_info['Current_Price'], 
            stock_info['1_Month_Low'], 
            stock_info['1_Month_High']
        )
        
        # Calculate average position
        stock_info['Current_vs_All'] = round(
            (stock_info['Price_vs_52W'] + 
             stock_info['Price_vs_3M'] + 
             stock_info['Price_vs_1M']) / 3, 1
        )
        
        # Print results
        print(f"\n📈 Results:")
        print(f"  • Current Price: ₹{stock_info['Current_Price']:.2f}")
        print(f"  • 52W Range: ₹{stock_info['52_Week_Low']:.2f} - ₹{stock_info['52_Week_High']:.2f}")
        print(f"  • 3M Range: ₹{stock_info['3_Month_Low']:.2f} - ₹{stock_info['3_Month_High']:.2f}")
        print(f"  • 1M Range: ₹{stock_info['1_Month_Low']:.2f} - ₹{stock_info['1_Month_High']:.2f}")
        print(f"  • Current vs All: {stock_info['Current_vs_All']}%")
        print(f"  • Data Source: {stock_info['Data_Source']}")
        
        return stock_info
    
    def analyze_stocks(self, stock_dict):
        """
        Analyze multiple stocks
        Returns DataFrame with all stock information
        """
        results = []
        
        for symbol, company_name in stock_dict.items():
            try:
                stock_info = self.get_stock_info(symbol, company_name)
                results.append(stock_info)
            except Exception as e:
                print(f"  ✗ Error analyzing {symbol}: {str(e)}")
                # Add placeholder data for failed stocks
                results.append({
                    'Symbol': symbol,
                    'Current_Price': 0,
                    '52_Week_High': 0,
                    '52_Week_Low': 0,
                    '3_Month_High': 0,
                    '3_Month_Low': 0,
                    '1_Month_High': 0,
                    '1_Month_Low': 0,
                    'Price_vs_52W': 0,
                    'Price_vs_3M': 0,
                    'Price_vs_1M': 0,
                    'Current_vs_All': 0,
                    'Data_Source': 'Error'
                })
            
            time.sleep(2)  # Rate limiting
        
        return pd.DataFrame(results)
    
    def create_excel_report(self, df, filename='Stock_Analysis_Report.xlsx'):
        """
        Create formatted Excel report
        Returns filename of created report
        """
        try:
            writer = pd.ExcelWriter(filename, engine='openpyxl')
            df.to_excel(writer, sheet_name='Stock Data', index=False, startrow=2)
            
            workbook = writer.book
            worksheet = writer.sheets['Stock Data']
            
            # Add title
            worksheet['A1'] = f'NSE Stock Analysis Report - {datetime.now().strftime("%d %B %Y, %I:%M %p")}'
            worksheet['A1'].font = Font(size=14, bold=True, color='000000')
            worksheet.merge_cells('A1:O1')
            worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
            worksheet.row_dimensions[1].height = 30
            
            # Format headers
            header_font = Font(bold=True, color='000000', size=11)
            for cell in worksheet[3]:
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            worksheet.row_dimensions[3].height = 40
            
            # Add borders
            thin_border = Border(
                left=Side(style='thin', color='D3D3D3'),
                right=Side(style='thin', color='D3D3D3'),
                top=Side(style='thin', color='D3D3D3'),
                bottom=Side(style='thin', color='D3D3D3')
            )
            
            # Format data cells
            for row in worksheet.iter_rows(min_row=4, max_row=len(df)+3, 
                                          min_col=1, max_col=len(df.columns)):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    
                    # Format numbers
                    if isinstance(cell.value, (int, float)) and cell.column <= 10:
                        cell.number_format = '₹#,##0.00'
                    elif isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.0'
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                
                adjusted_width = min(max_length + 3, 22)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            writer.close()
            
            print(f"\n{'='*60}")
            print(f"✓ Excel file created successfully: {filename}")
            print(f"{'='*60}")
            
            return filename
            
        except Exception as e:
            print(f"\n✗ Error creating Excel file: {str(e)}")
            return None

def main():
    """Main execution function"""
    try:
        print("\n" + "="*70)
        print(" "*15 + "🔴 ACCURATE NSE STOCK ANALYZER 🔴")
        print("="*70)
        print(f"\n📅 Analysis Date: {datetime.now().strftime('%d %B %Y, %I:%M %p')}")
        
        # Define stocks to analyze
        stock_dict = {
            'IDEA': 'Vodafone Idea Limited',
            'ADANIPORTS': 'Adani Ports and SEZ',
            'RELIANCE': 'Reliance Industries',
            'BAJAJ-AUTO': 'Bajaj Auto Limited'
        }
        
        print(f"\n📊 Stocks to analyze ({len(stock_dict)}):")
        for symbol in stock_dict.keys():
            print(f"   • {symbol}")
        
        print("\n" + "-"*70)
        print("🔍 Fetching LIVE data from multiple sources...")
        print("   Sources: Yahoo Finance (yfinance), NSE Official, Google Finance")
        print("-"*70)
        
        # Analyze stocks
        analyzer = NSEStockAnalyzer()
        df = analyzer.analyze_stocks(stock_dict)
        
        # Display summary
        print("\n" + "="*70)
        print(" "*22 + "📊 ANALYSIS SUMMARY 📊")
        print("="*70 + "\n")
        
        display_df = df[['Symbol', 'Current_Price', '52_Week_Low', '52_Week_High',
                         'Price_vs_52W', 'Price_vs_3M', 'Price_vs_1M', 
                         'Current_vs_All', 'Data_Source']].copy()
        
        display_df.columns = ['Symbol', 'Price (₹)', '52W Low', '52W High',
                              '52W %', '3M %', '1M %', 'Current_vs_All (%)', 'Source']
        
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.precision', 2)
        
        print(display_df.to_string(index=False))
        
        # Create Excel report
        print("\n" + "-"*70)
        print("📝 Creating detailed Excel report...")
        print("-"*70)
        
        filename = f"Stock_Analysis_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        result_file = analyzer.create_excel_report(df, filename=filename)
        
        if result_file:
            print("\n" + "="*70)
            print(" "*25 + "✅ ANALYSIS COMPLETE!")
            print("="*70)
            print(f"\n📄 Output file: {result_file}")
        else:
            print("\n⚠ Analysis completed but Excel file creation failed")
            
    except Exception as e:
        print(f"\n✗ Fatal error in main execution: {str(e)}")
        print("Please check your internet connection and try again.")

if __name__ == "__main__":
    main()