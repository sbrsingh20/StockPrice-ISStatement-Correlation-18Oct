import pandas as pd
import streamlit as st
import os

# Set up Streamlit app
st.title('Stock Analysis Based on Income Statement Data')

# Create a sidebar for user input
st.sidebar.header('Search for a Stock')
stock_name = st.sidebar.text_input('Enter Stock Symbol:', '')

# Function to fetch income statement data for the selected stock from the financials folder
def fetch_income_statement(stock_symbol):
    folder_path = 'financials'
    stock_data = None
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            # Read the Income Statement sheet
            try:
                income_statement = pd.read_excel(file_path, sheet_name='IncomeStatement')
                if 'Stock Symbol' in income_statement.columns and stock_symbol in income_statement['Stock Symbol'].values:
                    stock_data = income_statement[income_statement['Stock Symbol'] == stock_symbol]
                    return stock_data
            except Exception as e:
                st.warning(f"Error reading {filename}: {e}")
    return None

# Check if user has entered a stock symbol
if stock_name:
    additional_income_data = fetch_income_statement(stock_name)
    if additional_income_data is not None and not additional_income_data.empty:
        st.subheader(f'Income Statement for {stock_name}')
        st.write(additional_income_data)
    else:
        st.warning('No income statement data found for the selected stock.')
