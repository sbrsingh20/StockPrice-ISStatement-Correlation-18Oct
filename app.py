import pandas as pd
import streamlit as st
import os

# Load the data from Excel files
inflation_data = pd.read_excel('Inflation_event_stock_analysis_resultsOct.xlsx')
income_data = pd.read_excel('Inflation_IncomeStatement_correlation_results.xlsx')
interest_rate_data = pd.read_excel('interestrate_event_stock_analysis_resultsOct.xlsx')
interest_rate_income_data = pd.read_excel('interestrate_IncomeStatement_correlation_results.xlsx')

# Set up Streamlit app
st.title('Stock Analysis Based on Economic Events')

# Create a sidebar for user input
st.sidebar.header('Search for a Stock')
event_type = st.sidebar.selectbox('Select Event Type:', ['Inflation', 'Interest Rate'])
stock_name = st.sidebar.text_input('Enter Stock Symbol:', '')
expected_event_rate = st.sidebar.number_input('Enter Expected Upcoming Rate (%):', value=3.65, step=0.01)

# Function to fetch income statement data from the financials folder
def fetch_income_statement(stock_symbol):
    folder_path = 'financials'
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            # Read the Income Statement sheet
            try:
                income_statement = pd.read_excel(file_path, sheet_name='IncomeStatement')
                if stock_symbol in income_statement['Stock Symbol'].values:
                    # Filter for the specific stock
                    stock_data = income_statement[income_statement['Stock Symbol'] == stock_symbol]
                    return stock_data
            except Exception as e:
                st.warning(f"Error reading {filename}: {e}")
    return None

# Function to fetch details for a specific stock based on the event type
def get_stock_details(stock_symbol, event_type):
    if event_type == 'Inflation':
        event_row = inflation_data[inflation_data['Symbol'] == stock_symbol]
        income_row = income_data[income_data['Stock Name'] == stock_symbol]
    else:  # Interest Rate
        event_row = interest_rate_data[interest_rate_data['Symbol'] == stock_symbol]
        income_row = interest_rate_income_data[interest_rate_income_data['Stock Name'] == stock_symbol]

    if not event_row.empty and not income_row.empty:
        event_details = event_row.iloc[0]
        income_details = income_row.iloc[0]

        st.subheader(f'Details for {stock_symbol}')
        
        # Display event data
        st.write("### Event Data")
        st.write(event_row)

        # Display income statement data
        st.write("### Income Statement Data")
        st.write(income_row)

        # Fetch additional income statement data from the financials folder
        additional_income_data = fetch_income_statement(stock_symbol)
        if additional_income_data is not None:
            st.write("### Additional Income Statement Data")
            st.write(additional_income_data)

            # Project values based on expected event rate
            generate_projections(event_details, income_details, expected_event_rate, event_type, additional_income_data)
        
        # Interpret event and income data
        interpret_event_data(event_details, event_type)
        interpret_income_data(income_details)
    else:
        st.warning('Stock symbol not found in the data.')

# Function to generate projections based on expected event rate
def generate_projections(event_details, income_details, expected_event_rate, event_type, additional_income_data):
    latest_event_value = income_details['Latest Event Value']
    event_change = expected_event_rate - latest_event_value
    
    # Create a DataFrame to store the results
    projections = pd.DataFrame(columns=['Parameter', 'Current Value', 'Projected Value', 'Change'])

    # Check available columns in event details
    st.write("Available parameters in event details:", event_details.index)

    # Check if 'Latest Close Price' exists
    if 'Latest Close Price' in event_details.index:
        latest_close_price = pd.to_numeric(event_details['Latest Close Price'], errors='coerce')
        price_change = event_details['Event Coefficient'] * event_change
        projected_price = latest_close_price + price_change
        
        new_row = pd.DataFrame([{
            'Parameter': 'Projected Stock Price',
            'Current Value': latest_close_price,
            'Projected Value': projected_price,
            'Change': price_change
        }])
        projections = pd.concat([projections, new_row], ignore_index=True)
    else:
        st.warning("Stock Price data not available in event details.")

    # Project changes in additional income statement items
    for column in additional_income_data.columns:
        if column not in ['Date', 'Stock Symbol']:
            current_value = pd.to_numeric(additional_income_data[column].iloc[-1], errors='coerce')  # Get the last entry
            if pd.notna(current_value):
                projected_value = current_value * (1 + event_change / 100)  # Simplified assumption for projection
                change = projected_value - current_value
                
                new_row = pd.DataFrame([{
                    'Parameter': column,
                    'Current Value': current_value,
                    'Projected Value': projected_value,
                    'Change': change
                }])
                projections = pd.concat([projections, new_row], ignore_index=True)

    # Display the projections table
    st.write("### Projected Changes Based on Expected Event")
    st.dataframe(projections)

# Check if user has entered a stock symbol
if stock_name:
    get_stock_details(stock_name, event_type)
