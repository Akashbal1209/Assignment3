import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import os

class OptionsAnalyzer:
    def __init__(self):
        self.analysis_data = []
        self.margin_percent = 15  # Default margin percentage
        self.lot_multiplier = 1   # Default lot multiplier
        
    def calculate_percentile(self, current, high, low):
        if high == low:
            return 50.0
        return ((current - low) / (high - low)) * 100
    
    def find_nearest_strike(self, price, strike_interval=50):
        return round(price / strike_interval) * strike_interval
    
    def calculate_strike_with_margin(self, spot_price, margin_percent, is_call=True):
        nearest_strike = self.find_nearest_strike(spot_price)
        
        if is_call:
            target_price = nearest_strike * (1 - margin_percent / 100)
        else:
            target_price = nearest_strike * (1 + margin_percent / 100)
        
        return self.find_nearest_strike(target_price)
    
    def generate_option_premium(self, spot_price, strike_price, option_type):
        if option_type == 'CE':
            if spot_price > strike_price:
                base_premium = (spot_price - strike_price) + (spot_price * 0.02)
            else:
                base_premium = spot_price * 0.01
        else:  # PE
            if spot_price < strike_price:
                base_premium = (strike_price - spot_price) + (spot_price * 0.02)
            else:
                base_premium = spot_price * 0.01
        
        return round(base_premium, 2)
    
    def calculate_irr(self, premium, strike_price, days_to_expiry=30):
        margin_required = strike_price * 0.15
        if margin_required == 0:
            return 0
        irr = (premium / margin_required) * (365 / days_to_expiry) * 100
        return round(irr, 2)
    
    def get_sample_data(self):
        """Sample stock data - Replace with actual NSE API calls"""
        return [
            {'symbol': 'NIFTY', 'spot_price': 19500, 'high_52w': 20000, 'low_52w': 18000, 'lot_size': 50},
            {'symbol': 'BANKNIFTY', 'spot_price': 44500, 'high_52w': 46000, 'low_52w': 42000, 'lot_size': 25},
            {'symbol': 'RELIANCE', 'spot_price': 2450, 'high_52w': 2650, 'low_52w': 2250, 'lot_size': 250},
            {'symbol': 'TCS', 'spot_price': 3600, 'high_52w': 3850, 'low_52w': 3200, 'lot_size': 125},
            {'symbol': 'INFY', 'spot_price': 1480, 'high_52w': 1600, 'low_52w': 1350, 'lot_size': 300},
            {'symbol': 'HDFCBANK', 'spot_price': 1650, 'high_52w': 1750, 'low_52w': 1450, 'lot_size': 550},
            {'symbol': 'ICICIBANK', 'spot_price': 950, 'high_52w': 1050, 'low_52w': 850, 'lot_size': 1375},
            {'symbol': 'SBIN', 'spot_price': 590, 'high_52w': 650, 'low_52w': 520, 'lot_size': 1500},
            {'symbol': 'BHARTIARTL', 'spot_price': 880, 'high_52w': 950, 'low_52w': 750, 'lot_size': 1220},
            {'symbol': 'HINDUNILVR', 'spot_price': 2580, 'high_52w': 2800, 'low_52w': 2350, 'lot_size': 300}
        ]
    
    def analyze(self):
        """Main analysis function"""
        print("🔄 Starting Options Analysis...")
        
        # Get stock data
        stocks = self.get_sample_data()
        
        for stock in stocks:
            spot_price = stock['spot_price']
            high_52w = stock['high_52w']
            low_52w = stock['low_52w']
            
            # Calculate percentile
            percentile = self.calculate_percentile(spot_price, high_52w, low_52w)
            
            # Calculate strikes
            ce_strike = self.calculate_strike_with_margin(spot_price, self.margin_percent, is_call=True)
            pe_strike = self.calculate_strike_with_margin(spot_price, self.margin_percent, is_call=False)
            
            # Generate premiums
            ce_premium = self.generate_option_premium(spot_price, ce_strike, 'CE')
            pe_premium = self.generate_option_premium(spot_price, pe_strike, 'PE')
            
            # Calculate IRR
            ce_irr = self.calculate_irr(ce_premium, ce_strike)
            pe_irr = self.calculate_irr(pe_premium, pe_strike)
            
            # Adjusted lot size
            adjusted_lot_size = stock['lot_size'] * self.lot_multiplier
            
            # Store data
            row_data = {
                'Symbol': stock['symbol'],
                'Spot Price': spot_price,
                '52W High': high_52w,
                '52W Low': low_52w,
                'Percentile': round(percentile, 2),
                'Lot Size': adjusted_lot_size,
                'CE Strike': ce_strike,
                'CE Premium': ce_premium,
                'CE IRR': round(ce_irr, 2),
                'PE Strike': pe_strike,
                'PE Premium': pe_premium,
                'PE IRR': round(pe_irr, 2),
                'Margin Used (%)': self.margin_percent
            }
            
            self.analysis_data.append(row_data)
        
        print(f"✅ Analysis completed for {len(self.analysis_data)} symbols!")
        return self.analysis_data
    
    def save_to_excel(self, filename=None):
        """Save analysis to Excel file"""
        if not self.analysis_data:
            print("❌ No data to export. Run analyze() first.")
            return
        
        # Create DataFrame
        df = pd.DataFrame(self.analysis_data)
        
        # Generate filename if not provided
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Options_Analysis_{timestamp}.xlsx"
        
        # Export to Excel
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Options Analysis', index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets['Options Analysis']
            for idx, col in enumerate(df.columns):
                max_length = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 20)
        
        print(f"✅ Excel file saved: {filename}")
        return filename
    
    def save_graphs(self, filename=None):
        """Generate and save all graphs"""
        if not self.analysis_data:
            print("❌ No data to plot. Run analyze() first.")
            return
        
        print("📊 Generating graphs...")
        
        # Create DataFrame
        df = pd.DataFrame(self.analysis_data)
        
        # Generate filename if not provided
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Options_Analysis_Graphs_{timestamp}.png"
        
        # Create figure with subplots
        fig, axes = plt.subplots(2, 3, figsize=(16, 10))
        fig.suptitle('NSE Options Trading Analysis Dashboard', fontsize=18, fontweight='bold', y=0.995)
        
        x = range(len(df))
        width = 0.35
        
        # 1. IRR Comparison (CE vs PE)
        ax1 = axes[0, 0]
        ax1.bar([i - width/2 for i in x], df['CE IRR'], width, label='CE IRR', color='#667eea', alpha=0.8)
        ax1.bar([i + width/2 for i in x], df['PE IRR'], width, label='PE IRR', color='#f093fb', alpha=0.8)
        ax1.set_xlabel('Symbol', fontsize=10, fontweight='bold')
        ax1.set_ylabel('IRR (%)', fontsize=10, fontweight='bold')
        ax1.set_title('Call vs Put IRR Comparison', fontsize=12, fontweight='bold')
        ax1.set_xticks(x)
        ax1.set_xticklabels(df['Symbol'], rotation=45, ha='right', fontsize=9)
        ax1.legend(fontsize=9)
        ax1.grid(True, alpha=0.3)
        
        # 2. Premium Comparison
        ax2 = axes[0, 1]
        ax2.bar([i - width/2 for i in x], df['CE Premium'], width, label='CE Premium', color='#4facfe', alpha=0.8)
        ax2.bar([i + width/2 for i in x], df['PE Premium'], width, label='PE Premium', color='#00f2fe', alpha=0.8)
        ax2.set_xlabel('Symbol', fontsize=10, fontweight='bold')
        ax2.set_ylabel('Premium (₹)', fontsize=10, fontweight='bold')
        ax2.set_title('Call vs Put Premium Comparison', fontsize=12, fontweight='bold')
        ax2.set_xticks(x)
        ax2.set_xticklabels(df['Symbol'], rotation=45, ha='right', fontsize=9)
        ax2.legend(fontsize=9)
        ax2.grid(True, alpha=0.3)
        
        # 3. 52-Week Percentile
        ax3 = axes[0, 2]
        colors = ['#43e97b' if p >= 50 else '#fa709a' for p in df['Percentile']]
        bars = ax3.barh(df['Symbol'], df['Percentile'], color=colors, alpha=0.8)
        ax3.set_xlabel('Percentile (%)', fontsize=10, fontweight='bold')
        ax3.set_title('52-Week Price Percentile', fontsize=12, fontweight='bold')
        ax3.axvline(x=50, color='red', linestyle='--', linewidth=2, label='50% Mark')
        ax3.legend(fontsize=9)
        ax3.grid(True, alpha=0.3, axis='x')
        
        # Add value labels
        for i, bar in enumerate(bars):
            width_val = bar.get_width()
            ax3.text(width_val + 1, bar.get_y() + bar.get_height()/2, 
                    f'{width_val:.1f}%', va='center', fontsize=8)
        
        # 4. Spot Price vs Strikes
        ax4 = axes[1, 0]
        ax4.plot(df['Symbol'], df['Spot Price'], marker='o', label='Spot Price', 
                linewidth=2.5, markersize=8, color='#667eea')
        ax4.plot(df['Symbol'], df['CE Strike'], marker='s', label='CE Strike', 
                linewidth=2, markersize=6, color='#4facfe', linestyle='--')
        ax4.plot(df['Symbol'], df['PE Strike'], marker='^', label='PE Strike', 
                linewidth=2, markersize=6, color='#f093fb', linestyle='--')
        ax4.set_xlabel('Symbol', fontsize=10, fontweight='bold')
        ax4.set_ylabel('Price (₹)', fontsize=10, fontweight='bold')
        ax4.set_title('Spot Price vs Strike Prices', fontsize=12, fontweight='bold')
        ax4.legend(fontsize=9)
        ax4.grid(True, alpha=0.3)
        plt.setp(ax4.xaxis.get_majorticklabels(), rotation=45, ha='right', fontsize=9)
        
        # 5. Average IRR by Type
        ax5 = axes[1, 1]
        avg_ce_irr = df['CE IRR'].mean()
        avg_pe_irr = df['PE IRR'].mean()
        bars = ax5.bar(['Call Options', 'Put Options'], [avg_ce_irr, avg_pe_irr], 
               color=['#667eea', '#f093fb'], alpha=0.8, width=0.6)
        ax5.set_ylabel('Average IRR (%)', fontsize=10, fontweight='bold')
        ax5.set_title('Average IRR Comparison', fontsize=12, fontweight='bold')
        ax5.grid(True, alpha=0.3, axis='y')
        
        # Add value labels
        for bar in bars:
            height = bar.get_height()
            ax5.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f'{height:.2f}%', ha='center', va='bottom', fontweight='bold', fontsize=10)
        
        # 6. Premium to Strike Ratio
        ax6 = axes[1, 2]
        df['CE_Ratio'] = (df['CE Premium'] / df['CE Strike']) * 100
        df['PE_Ratio'] = (df['PE Premium'] / df['PE Strike']) * 100
        ax6.scatter(df['Symbol'], df['CE_Ratio'], s=120, alpha=0.7, 
                   label='CE Ratio', color='#667eea', edgecolors='black', linewidth=0.5)
        ax6.scatter(df['Symbol'], df['PE_Ratio'], s=120, alpha=0.7, 
                   label='PE Ratio', color='#f093fb', edgecolors='black', linewidth=0.5)
        ax6.set_xlabel('Symbol', fontsize=10, fontweight='bold')
        ax6.set_ylabel('Premium/Strike Ratio (%)', fontsize=10, fontweight='bold')
        ax6.set_title('Premium to Strike Price Ratio', fontsize=12, fontweight='bold')
        ax6.legend(fontsize=9)
        ax6.grid(True, alpha=0.3)
        plt.setp(ax6.xaxis.get_majorticklabels(), rotation=45, ha='right', fontsize=9)
        
        plt.tight_layout()
        
        # Save to file
        fig.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close(fig)
        
        print(f"✅ Graphs saved: {filename}")
        return filename
    
    def run(self, margin_percent=15, lot_multiplier=1):
        """Run complete analysis and save outputs"""
        print("\n" + "="*60)
        print("🚀 NSE OPTIONS TRADING ANALYZER")
        print("="*60)
        
        # Set parameters
        self.margin_percent = margin_percent
        self.lot_multiplier = lot_multiplier
        
        print(f"\n⚙️  Configuration:")
        print(f"   - Margin Percentage: {margin_percent}%")
        print(f"   - Lot Size Multiplier: {lot_multiplier}x")
        print()
        
        # Run analysis
        self.analyze()
        
        # Save to Excel
        excel_file = self.save_to_excel()
        
        # Save graphs
        graph_file = self.save_graphs()
        
        print("\n" + "="*60)
        print("✨ ANALYSIS COMPLETE!")
        print("="*60)
        print(f"📄 Excel Report: {excel_file}")
        print(f"📊 Graphs: {graph_file}")
        print("="*60 + "\n")
        
        return excel_file, graph_file


# Main execution
if __name__ == "__main__":
    analyzer = OptionsAnalyzer()
    
    analyzer.run(margin_percent=15, lot_multiplier=1)

