import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io
import os

class DCFAnalysis:
    def __init__(self):
        self.statements = None
        self.projections = None
        self.wacc = None
        self.terminal_growth = None
        self.projection_years = None
        
    def load_data(self, filepath):
        """Load financial statements from CSV file"""
        try:
            self.statements = pd.read_csv(filepath)
            required_columns = ['Revenue', 'Operating_Expenses', 'Depreciation', 'Capex', 'Change_Working_Capital']
            missing_columns = [col for col in required_columns if col not in self.statements.columns]
            if missing_columns:
                raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
            print("Financial statements loaded successfully!")
        except Exception as e:
            print(f"Error loading file: {e}")
            return False
        return True

    def get_user_inputs(self):
        """Gather necessary inputs from user"""
        print("\n=== DCF Analysis Parameters ===")
        
        # Get WACC
        while True:
            try:
                self.wacc = float(input("Enter WACC (as decimal, e.g., 0.10 for 10%): "))
                if 0 < self.wacc < 1:
                    break
                print("WACC should be between 0 and 1")
            except ValueError:
                print("Please enter a valid number")

        # Get terminal growth rate
        while True:
            try:
                self.terminal_growth = float(input("Enter terminal growth rate (as decimal): "))
                if -0.1 < self.terminal_growth < 0.1:
                    break
                print("Terminal growth rate should be between -0.1 and 0.1")
            except ValueError:
                print("Please enter a valid number")

        # Get projection years
        while True:
            try:
                self.projection_years = int(input("Enter number of years to project (1-10): "))
                if 1 <= self.projection_years <= 10:
                    break
                print("Projection years should be between 1 and 10")
            except ValueError:
                print("Please enter a valid number")

    def calculate_free_cash_flow(self):
        """Calculate historical free cash flow"""
        try:
            self.statements['EBIT'] = self.statements['Revenue'] - self.statements['Operating_Expenses']
            self.statements['NOPAT'] = self.statements['EBIT'] * (1 - 0.21)  # Assuming 21% tax rate
            self.statements['FCF'] = (self.statements['NOPAT'] + 
                                    self.statements['Depreciation'] - 
                                    self.statements['Capex'] - 
                                    self.statements['Change_Working_Capital'])
            return self.statements['FCF']
        except Exception as e:
            print(f"Error calculating free cash flow: {e}")
            return None

    def project_financials(self):
        """Project future cash flows"""
        try:
            # Calculate growth rates
            historical_growth = self.statements['Revenue'].pct_change().mean()
            
            # Create projection DataFrame
            years = range(1, self.projection_years + 1)
            self.projections = pd.DataFrame(index=years)
            
            # Project revenue with declining growth
            growth_rates = np.linspace(historical_growth, self.terminal_growth, self.projection_years)
            base_revenue = self.statements['Revenue'].iloc[-1]
            
            self.projections['Revenue'] = [base_revenue * np.prod(1 + growth_rates[:i]) 
                                         for i in range(1, self.projection_years + 1)]
            
            # Project other metrics based on historical ratios
            avg_margin = (self.statements['EBIT'] / self.statements['Revenue']).mean()
            self.projections['EBIT'] = self.projections['Revenue'] * avg_margin
            self.projections['FCF'] = self.projections['EBIT'] * (1 - 0.21)  # Simplified FCF projection
            
            # Convert projections to regular float values
            self.projections = self.projections.astype(float)
            
        except Exception as e:
            print(f"Error projecting financials: {e}")
            return None

    def calculate_dcf_value(self):
        """Calculate DCF value"""
        try:
            # Calculate present value of projected FCF
            pv_fcf = sum(self.projections['FCF'] / (1 + self.wacc) ** year 
                         for year in range(1, self.projection_years + 1))
            
            # Calculate terminal value
            terminal_fcf = self.projections['FCF'].iloc[-1] * (1 + self.terminal_growth)
            terminal_value = terminal_fcf / (self.wacc - self.terminal_growth)
            pv_terminal = terminal_value / (1 + self.wacc) ** self.projection_years
            
            # Total value
            total_value = pv_fcf + pv_terminal
            
            # Convert all values to regular floats
            return {
                'PV of FCF': float(pv_fcf),
                'Terminal Value': float(terminal_value),
                'PV of Terminal Value': float(pv_terminal),
                'Total Value': float(total_value)
            }
        except Exception as e:
            print(f"Error calculating DCF value: {e}")
            return None

    def create_visualizations(self):
        """Create visualizations for the analysis"""
        try:
            # Set style
            sns.set_style('darkgrid')
            plt.style.use('seaborn')
            figures = []

            # Historical vs Projected FCF
            fig1, ax1 = plt.subplots(figsize=(10, 6))
            historical_years = range(-len(self.statements), 0)
            projected_years = range(1, self.projection_years + 1)
            
            ax1.plot(historical_years, self.statements['FCF'], 'b-', label='Historical FCF')
            ax1.plot(projected_years, self.projections['FCF'], 'r--', label='Projected FCF')
            ax1.set_title('Free Cash Flow - Historical vs Projected')
            ax1.set_xlabel('Year')
            ax1.set_ylabel('FCF ($)')
            ax1.legend()
            figures.append(fig1)

            # Revenue Growth Rates
            fig2, ax2 = plt.subplots(figsize=(10, 6))
            growth_rates = np.linspace(
                self.statements['Revenue'].pct_change().mean(),
                self.terminal_growth,
                self.projection_years
            )
            ax2.plot(projected_years, growth_rates * 100, 'g-')
            ax2.set_title('Projected Revenue Growth Rates')
            ax2.set_xlabel('Year')
            ax2.set_ylabel('Growth Rate (%)')
            figures.append(fig2)

            return figures
        except Exception as e:
            print(f"Error creating visualizations: {e}")
            return None

    def generate_pdf_report(self, figures, dcf_value, output_filename='dcf_analysis.pdf'):
        """Generate PDF report with analysis results"""
        try:
            doc = SimpleDocTemplate(output_filename, pagesize=letter)
            styles = getSampleStyleSheet()
            story = []

            # Title
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=24,
                spaceAfter=30
            )
            story.append(Paragraph("DCF Analysis Report", title_style))
            story.append(Spacer(1, 20))

            # Key Assumptions
            story.append(Paragraph("Key Assumptions:", styles['Heading2']))
            assumptions = [
                f"WACC: {self.wacc:.1%}",
                f"Terminal Growth Rate: {self.terminal_growth:.1%}",
                f"Projection Years: {self.projection_years}"
            ]
            for assumption in assumptions:
                story.append(Paragraph(assumption, styles['Normal']))
            story.append(Spacer(1, 20))

            # DCF Results
            story.append(Paragraph("Valuation Results:", styles['Heading2']))
            # Ensure all values are regular floats and format them
            results_data = [[k, f"${v:,.2f}"] for k, v in dcf_value.items()]
            results_table = Table(results_data)
            results_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                ('TOPPADDING', (0, 0), (-1, -1), 12),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            story.append(results_table)
            story.append(Spacer(1, 20))

            # Add figures
            for i, fig in enumerate(figures):
                imgdata = io.BytesIO()
                fig.savefig(imgdata, format='png', bbox_inches='tight')
                imgdata.seek(0)
                img = Image(imgdata)
                img.drawHeight = 300
                img.drawWidth = 500
                story.append(img)
                story.append(Spacer(1, 20))

            # Build PDF
            doc.build(story)
            print(f"\nPDF report generated: {output_filename}")
            return True
        except Exception as e:
            print(f"Error generating PDF report: {e}")
            print(f"DCF Values types:", {k: type(v) for k, v in dcf_value.items()})
            return False

def create_test_data(filename='test_financial_data.csv'):
    """Create synthetic financial data for testing"""
    try:
        # Create 5 years of historical data
        years = range(2020, 2025)
        
        # Generate synthetic data with realistic growth and patterns
        data = {
            'Year': list(years),
            'Revenue': [1000000 * (1.1 ** i) for i in range(5)],  # 10% annual growth
            'Operating_Expenses': [800000 * (1.08 ** i) for i in range(5)],  # 8% annual growth
            'Depreciation': [50000 * (1.05 ** i) for i in range(5)],  # 5% annual growth
            'Capex': [70000 * (1.05 ** i) for i in range(5)],  # 5% annual growth
            'Change_Working_Capital': [20000 * (1.03 ** i) for i in range(5)]  # 3% annual growth
        }
        
        # Create DataFrame and save to CSV
        df = pd.DataFrame(data)
        df.to_csv(filename, index=False)
        print(f"Test data created successfully: {filename}")
        return filename
    except Exception as e:
        print(f"Error creating test data: {e}")
        return None

def test_run():
    """Perform a test run with synthetic data"""
    print("Starting DCF Analysis test run...")
    
    # Create test data
    test_file = create_test_data()
    if not test_file:
        return
    
    # Initialize DCF analysis
    dcf = DCFAnalysis()
    
    # Load test data
    if not dcf.load_data(test_file):
        return
    
    # Set test inputs
    dcf.wacc = 0.10
    dcf.terminal_growth = 0.02
    dcf.projection_years = 5
    
    print("\nTest inputs:")
    print(f"WACC: {dcf.wacc}")
    print(f"Terminal Growth: {dcf.terminal_growth}")
    print(f"Projection Years: {dcf.projection_years}")
    
    # Perform analysis
    fcf = dcf.calculate_free_cash_flow()
    if fcf is None:
        return
    
    dcf.project_financials()
    dcf_value = dcf.calculate_dcf_value()
    if dcf_value is None:
        return
    
    print("\nDCF Values:")
    for k, v in dcf_value.items():
        print(f"{k}: ${v:,.2f}")
    
    # Create visualizations
    figures = dcf.create_visualizations()
    if figures is None:
        return
    
    # Generate PDF report
    if not dcf.generate_pdf_report(figures, dcf_value, 'test_dcf_analysis.pdf'):
        return
    
    print("\nTest run completed successfully!")
    print("Please check 'test_dcf_analysis.pdf' for the results.")
    
    # Clean up test data file
    try:
        os.remove(test_file)
        print("Test data file cleaned up.")
    except Exception as e:
        print(f"Error cleaning up test file: {e}")

if __name__ == "__main__":
    test_run()