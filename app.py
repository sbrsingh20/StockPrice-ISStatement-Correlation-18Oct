import pandas as pd
import streamlit as st

# Load the data from Excel files
inflation_data = pd.read_excel('Inflation_event_stock_analysis_resultsOct.xlsx')
income_data = pd.read_excel('Inflation_IncomeStatement_correlation_results.xlsx')

# Set up Streamlit app
st.title('Stock Analysis Based on Inflation Events')

# Create a sidebar for user input
st.sidebar.header('Search for a Stock')
stock_name = st.sidebar.text_input('Enter Stock Symbol:', '')

# User input for expected upcoming inflation
expected_inflation = st.sidebar.number_input('Enter Expected Upcoming Inflation Rate (%):', value=3.65, step=0.01)

# Function to fetch details for a specific stock
def get_stock_details(stock_symbol):
    inflation_row = inflation_data[inflation_data['Symbol'] == stock_symbol]
    income_row = income_data[income_data['Stock Name'] == stock_symbol]

    if not inflation_row.empty and not income_row.empty:
        inflation_details = inflation_row.iloc[0]
        income_details = income_row.iloc[0]

        st.subheader(f'Details for {stock_symbol}')
        
        # Display inflation event data
        st.write("### Inflation Event Data")
        st.write(inflation_row)

        # Display income statement data
        st.write("### Income Statement Data")
        st.write(income_row)

        # Generate projections based on expected inflation
        generate_projections(inflation_details, income_details, expected_inflation)
        
        # Additional interpretations based on conditions
        interpret_inflation_data(inflation_details)
        interpret_income_data(income_details)
    else:
        st.warning('Stock symbol not found in the data.')

# Function to interpret inflation data
def interpret_inflation_data(details):
    st.write("### Interpretation of Inflation Event Data")
    if details['Event Coefficient'] < -1:
        st.write("**1% Increase in Inflation:** Stock price decreases significantly. Increase portfolio risk.")
    elif details['Event Coefficient'] > 1:
        st.write("**1% Increase in Inflation:** Stock price increases, benefiting from inflation.")

# Function to interpret income data
def interpret_income_data(details):
    st.write("### Interpretation of Income Statement Data")
    if 'Average Operating Margin' in details.index:
        average_operating_margin = details['Average Operating Margin']
        if average_operating_margin > 0.2:
            st.write("**High Operating Margin:** Indicates strong management effectiveness.")
        elif average_operating_margin < 0.1:
            st.write("**Low Operating Margin:** Reflects risk in profitability.")

# Function to generate projections based on expected inflation
def generate_projections(inflation_details, income_details, expected_inflation):
    latest_event_value = income_details['Latest Event Value']  # Now getting the actual inflation value from Income Statement Data
    inflation_change = expected_inflation - latest_event_value
    
    # Create a DataFrame to store the results
    projections = pd.DataFrame(columns=['Parameter', 'Current Value', 'Projected Value', 'Change'])

    # Check available columns in inflation details
    st.write("Available parameters in inflation details:", inflation_details.index)

    # Check if 'Latest Close Price' exists
    if 'Latest Close Price' in inflation_details.index:
        latest_close_price = pd.to_numeric(inflation_details['Latest Close Price'], errors='coerce')
        price_change = inflation_details['Event Coefficient'] * inflation_change
        projected_price = latest_close_price + price_change
        
        new_row = pd.DataFrame([{
            'Parameter': 'Projected Stock Price',
            'Current Value': latest_close_price,
            'Projected Value': projected_price,
            'Change': price_change
        }])
        projections = pd.concat([projections, new_row], ignore_index=True)
    else:
        st.warning("Stock Price data not available in inflation details.")

    # Project changes in income statement items
    for column in income_details.index:
        if column != 'Stock Name':
            current_value = pd.to_numeric(income_details[column], errors='coerce')
            if pd.notna(current_value):  # Check if the conversion was successful
                if column in inflation_details.index:  # Check if there is a correlation factor
                    correlation_factor = inflation_details[column] if column in inflation_details.index else 0
                    projected_value = current_value + (current_value * correlation_factor * inflation_change / 100)
                else:
                    projected_value = current_value * (1 + inflation_change / 100)  # Simplified assumption
                
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
    st.write("### Projected Changes Based on Expected Inflation")
    st.dataframe(projections)

# Check if user has entered a stock symbol
if stock_name:
    get_stock_details(stock_name)
