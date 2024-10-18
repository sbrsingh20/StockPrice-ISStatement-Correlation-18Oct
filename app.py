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

        # Fetch additional income statement data for the selected stock
        additional_income_data = fetch_income_statement(stock_symbol)
        if additional_income_data is not None and not additional_income_data.empty:
            st.write("### Additional Income Statement Data")
            st.write(additional_income_data)

            # Generate projections based on expected event rate
            generate_projections(event_details, income_details, expected_event_rate, event_type, additional_income_data)
        
            # Interpret event and income data
            if 'Event Coefficient' in event_details.index:  # Check if the column exists
                interpret_event_data(event_details, event_type)
            else:
                st.warning('Event Coefficient not found in event data.')
            
            interpret_income_data(income_details)
        else:
            st.warning('Additional income statement data not found.')
    else:
        st.warning('Stock symbol not found in the data.')

# Function to interpret event data
def interpret_event_data(details, event_type):
    st.write("### Interpretation of Event Data")
    if event_type == 'Inflation':
        if details.get('Event Coefficient', 0) < -1:  # Use .get to avoid KeyError
            st.write("**1% Increase in Inflation:** Stock price decreases significantly. Increase portfolio risk.")
        elif details.get('Event Coefficient', 0) > 1:
            st.write("**1% Increase in Inflation:** Stock price increases, benefiting from inflation.")
    else:  # Interest Rate
        if details.get('Event Coefficient', 0) < -1:
            st.write("**1% Increase in Interest Rate:** Stock price decreases significantly. Increase portfolio risk.")
        elif details.get('Event Coefficient', 0) > 1:
            st.write("**1% Increase in Interest Rate:** Stock price increases, benefiting from interest hikes.")

# Function to interpret income data
def interpret_income_data(details):
    st.write("### Interpretation of Income Statement Data")
    if 'Average Operating Margin' in details.index:
        average_operating_margin = details['Average Operating Margin']
        if average_operating_margin > 0.2:
            st.write("**High Operating Margin:** Indicates strong management effectiveness.")
        elif average_operating_margin < 0.1:
            st.write("**Low Operating Margin:** Reflects risk in profitability.")

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

    # Project changes in income statement items for the selected stock
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
