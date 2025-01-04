import streamlit as st
import pandas as pd

# Load stock data from the uploaded Excel file
file_path = "StockLexWarehouse.xlsx"  # Replace with your actual file path or file uploader
sheet_name = "รายการสินค้า"

# Load data into a DataFrame
stock_data = pd.read_excel(file_path, sheet_name=sheet_name)

# Convert column names to English for simplicity
stock_data.columns = [
    "Item Code", "Description", "Old Quantity", "Old Cost Price",
    "Old Selling Price", "Old Contractor Price", "New Quantity",
    "New Cost Price", "New Selling Price", "New Contractor Price"
]

# Initialize Streamlit app
st.title("Stock Management System")

# Display stock data
st.subheader("Current Stock Data")
if not stock_data.empty:
    st.write(stock_data)
else:
    st.info("No stock data available.")

# Select an item to Check
st.subheader("Check Stock Data")
item_code = st.selectbox("Select an item to Check:", stock_data["Description"].unique())

# Display current data for the selected item
selected_item = stock_data[stock_data["Description"] == item_code].iloc[0]
st.write("**Current Data:**")
st.table(selected_item)










