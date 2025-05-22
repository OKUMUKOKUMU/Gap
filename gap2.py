import streamlit as st
from datetime import datetime
import io

def create_html_report(data):
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Weekly Sales Report - Week {data['week_number']}</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; }}
            h1 {{ color: #2e86c1; text-align: center; }}
            h2 {{ color: #1a5276; border-bottom: 1px solid #ddd; padding-bottom: 5px; }}
            .negative {{ color: #e74c3c; }}
            .positive {{ color: #27ae60; }}
            ul {{ margin-top: 5px; }}
        </style>
    </head>
    <body>
        <h1>Weekly Sales Report – Week {data['week_number']} ({data['date_range']})</h1>
        
        <h2>Total Portfolio Sales</h2>
        <ul>
            <li>KSH {data['total_sales']:,.0f}, marking a 
                <span class={'negative' if data['total_change_percent'] < 0 else 'positive'}>
                {data['total_change_percent']}% change
                </span> from last week's KSH {data['last_week_sales']:,.0f}.</li>
            <li>{data['total_sales_comment']}</li>
        </ul>
        
        <h2>Performance by Sales Executives</h2>
        {''.join([f"""
        <h3>{exec_data['name']}</h3>
        <ul>
            <li><span class={'negative' if exec_data['change_percent'] < 0 else 'positive'}>
                {exec_data['change_percent']}%</span>: 
                KSH {exec_data['current_sales']/1000:.2f}K (from KSH {exec_data['last_sales']/1000:.2f}K)</li>
            <li>{exec_data['comment']}</li>
        </ul>
        """ for exec_data in data['executives']])}
        
        <h2>Key Highlights</h2>
        <ul>
            {''.join(f'<li>{highlight}</li>' for highlight in data['highlights'])}
        </ul>
        
        <h2>Strategic Next Steps</h2>
        <ol>
            {''.join(f'<li>{step}</li>' for step in data['next_steps'])}
        </ol>
    </body>
    </html>
    """
    return html

def main():
    st.title("Weekly Sales Report Generator (HTML Output)")
    
    # [Keep your existing input form code...]
    # Replace the Word generation part with:
    
    if submitted:
        report_data = {
            # [Keep your existing data structure...]
        }
        
        html_report = create_html_report(report_data)
        
        st.success("HTML report generated!")
        st.download_button(
            label="Download HTML Report",
            data=html_report,
            file_name=f"weekly_sales_report_week_{week_number}.html",
            mime="text/html"
        )

if __name__ == "__main__":
    main()