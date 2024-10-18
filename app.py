import pandas as pd
import streamlit as st

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

        # Generate projections based on expected event rate
        generate_projections(event_details, income_details, expected_event_rate, event_type)
        
        # Interpret event and income data
        interpret_event_data(event_details, event_type)
        interpret_income_data(income_details)
    else:
        st.warning('Stock symbol not found in the data.')

# Function to interpret event data
def interpret_event_data(details, event_type):
    st.write("### Interpretation of Event Data")
    if event_type == 'Inflation':
        if details['Event Coefficient'] < -1:
            st.write("**1% Increase in Inflation:** Stock price decreases significantly. Increase portfolio risk.")
        elif details['Event Coefficient'] > 1:
            st.write("**1% Increase in Inflation:** Stock price increases, benefiting from inflation.")
    else:  # Interest Rate
        if details['Event Coefficient'] < -1:
            st.write("**1% Increase in Interest Rate:** Stock price decreases significantly. Increase portfolio risk.")
        elif details['Event Coefficient'] > 1:
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
def generate_projections(event_details, income_details, expected_event_rate, event_type):
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

    # Project changes in income statement items
    for column in income_details.index:
        if column != 'Stock Name':
            current_value = pd.to_numeric(income_details[column], errors='coerce')
            if pd.notna(current_value):
                if column in event_details.index:
                    correlation_factor = event_details[column] if column in event_details.index else 0
                    projected_value = current_value + (current_value * correlation_factor * event_change / 100)
                else:
                    projected_value = current_value * (1 + event_change / 100)  # Simplified assumption
                
                change = projected_value - current_value
                
                new_row = pd.DataFrame([{
                    'Parameter': column,
                    'Current Value': current_value,
                    'Projected Value': projected_value,
                    'Change': change
                }])
                projections = pd.concat([projections, new_row], ignore_index=True)
            else:
                st.warning(f"Could not convert current value for {column} to numeric.")

    # Display the projections table
    st.write("### Projected Changes Based on Expected Event")
    st.dataframe(projections)

# Check if user has entered a stock symbol
if stock_name:
    get_stock_details(stock_name, event_type)
