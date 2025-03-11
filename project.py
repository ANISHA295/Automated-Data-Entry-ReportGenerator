import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import seaborn as sns
import pdfkit
import os

def read_data(file_path, sheet_name):
    """Read data from an Excel file"""
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df

def process_data(df):
    """Generate summary statistics"""
    summary = df.describe()
    return summary

def generate_chart(df, column_name, output_path):
    """Generate a bar chart for a specific column"""
    plt.figure(figsize=(10, 6))
    sns.countplot(data=df, x=column_name, palette='coolwarm')
    plt.title(f'Distribution of {column_name}')
    plt.xlabel(column_name)
    plt.ylabel('Count')
    plt.xticks(rotation=45)
    plt.savefig(output_path)
    plt.close()

def save_report(summary, output_file):
    """Save summary statistics to an Excel file"""
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='Summary')
    print(f"Report saved to {output_file}")

def generate_html_report(summary, chart_path, output_html, output_pdf):
    """Generate an HTML report with summary statistics and chart"""
    html_content = f"""
    <html>
    <head><title>Data Report</title></head>
    <body>
        <h1>Data Summary</h1>
        {summary.to_html()}
        <h2>Chart</h2>
        <img src='{chart_path}' alt='Chart'>
    </body>
    </html>
    """
    with open(output_html, "w", encoding='utf-8') as f:
        f.write(html_content)
    pdfkit.from_file(output_html, output_pdf)
    print(f"HTML and PDF reports saved as {output_html} and {output_pdf}")

def clean_old_reports(files):
    """Remove old report files before generating new ones"""
    for file in files:
        if os.path.exists(file):
            os.remove(file)
            print(f"Old file removed: {file}")
        
if __name__ == "__main__":
    input_file = "data.xlsx"  # Input Excel file
    sheet = "Sheet1"  # Sheet name
    output_report = "report.xlsx"  # Output report file
    chart_output = "chart.png"  # Chart output file
    column_to_chart = "Category"  # Column for visualization
    output_html = "report.html"  # HTML report
    output_pdf = "report.pdf"  # PDF report
    
    # Clean old reports before generating new ones
    clean_old_reports([output_report, chart_output, output_html, output_pdf])
    
    df = read_data(input_file, sheet)
    summary = process_data(df)
    save_report(summary, output_report)
    generate_chart(df, column_to_chart, chart_output)
    generate_html_report(summary, chart_output, output_html, output_pdf)
    
    print("Data processing, report generation, and visualization completed.")
