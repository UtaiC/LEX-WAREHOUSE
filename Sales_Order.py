######## Library ##########################################
import streamlit as st
import pandas as pd
from PIL import Image
import numpy as np
import os
import glob
from datetime import datetime, timedelta
import datetime
import calendar
import plotly.graph_objects as go
import requests
import sys
import streamlit.components.v1 as components
#
Logo=Image.open('WATTANA-Logo-Sales.jpg')
st.image(Logo,width=700)
## Table #####################################################
# # file='StockLexWarehouse.xlsx'
# file='https://docs.google.com/spreadsheets/d/10XwWMOqFxrQMiz_d7Vv__806167AQ8Wy6v82z9MOsQc/export?format=xlsx'
# df = pd.read_excel(file, header=0, engine='openpyxl')
# # df
Sales_Files='https://docs.google.com/spreadsheets/d/1tNRSxI9xv0FJMHbmwLpszP-ObAgpb0dY5AIciPmZWWY/export?format=xlsx'
Sales = pd.read_excel(Sales_Files, header=0, engine='openpyxl')
Sales['รายการสินค้า']=Sales['รายการสินค้า'].str.split(' / ').str[0]
Sales=pd.merge(Sales,df,left_on='รายการสินค้า',right_on='รายการ',how='left')
Sales.index = Sales.index + 1
###################################
F1,F2=st.columns([1,2])
with F1:

    Sales_No = st.selectbox('Sales_No',['WH-001', 'WH-002', 'WH-003', 'WH-004', 'WH-005', 'WH-006', 'WH-007', 'WH-008', 'WH-009', 'WH-010', 'WH-011', 'WH-012', 'WH-013', 'WH-014', 'WH-015', 'WH-016', 'WH-017', 'WH-018', 'WH-019', 'WH-020', 'WH-021', 'WH-022', 'WH-023', 'WH-024', 'WH-025', 'WH-026', 'WH-027', 'WH-028', 'WH-029', 'WH-030', 'WH-031', 'WH-032', 'WH-033', 'WH-034', 'WH-035', 'WH-036', 'WH-037', 'WH-038', 'WH-039', 'WH-040', 'WH-041', 'WH-042', 'WH-043', 'WH-044', 'WH-045', 'WH-046', 'WH-047', 'WH-048', 'WH-049', 'WH-050'] )
with F2:
    Sales_Date = st.selectbox('Sales_Date',['2024-01',	'2024-02',	'2024-03',	'2024-04',	'2024-05',	'2024-06',	'2024-07',	'2024-08',	'2024-09',	'2024-10',	'2024-11',	'2024-12'] )

Sales=pd.read_excel(Sales_No+'Sales-Update.xlsx')
# Stock Remain #################
Sales['จำนวนเก่า'] = Sales['จำนวนเก่า'].apply(pd.to_numeric, errors='coerce')
Sales['จำนวนใหม่'] = Sales['จำนวนใหม่'].apply(pd.to_numeric, errors='coerce')

# Calculate AMT
Sales['ราคาขายรวม'] = np.where(
    Sales['จำนวนเก่า'] < Sales["จำนวนสินค้า"],
    Sales['ราคาขายทุนเก่า'] * Sales['จำนวนเก่า'] + 
    Sales['ราคาขายทุนใหม่'] * (Sales["จำนวนสินค้า"] - Sales['จำนวนเก่า']),
    Sales['ราคาขายทุนเก่า'] * Sales["จำนวนสินค้า"])


Sales['ราคาขาย']=Sales['ราคาขายรวม']/Sales["จำนวนสินค้า"]

#######################
Sales=Sales[Sales["เลขที่เสนอราคา"].str.contains(Sales_No)]

Sales["Timestamp"] = pd.to_datetime(Sales["Timestamp"], errors='coerce')
Sales = Sales.dropna(subset=["Timestamp"])
Sales = Sales[Sales["Timestamp"].dt.strftime('%Y-%m-%d').str.startswith(Sales_Date)]

############################
# Calculate total available stock
Sales['ยอดรวมสต๊อก'] = Sales['จำนวนเก่า'] + Sales['จำนวนใหม่']

# Check for stock shortages
exceeds_stock = Sales["จำนวนสินค้า"] > Sales['ยอดรวมสต๊อก']

# If there are items with stock shortages, display them
if exceeds_stock.any():
    insufficient_stock_items = Sales[exceeds_stock]
    
    for index, row in insufficient_stock_items.iterrows():
        st.error(
            f'สต๊อกคงเหลือน้อยกว่ายอดสั่งชื้อสำหรับ: {row["รายการสินค้า"]} '
            f'(เหลือในสต๊อก: {row["ยอดรวมสต๊อก"]} ชิ้น)'
        )
############################
Date=Sales['Timestamp'].mean()
formatted_date = Date.strftime('%d-%m-%Y')  # or '%d/%m/%Y' for a different separator
# Cust = Sales[Sales['ลูกค้า'].notna()]
# Cust['ลูกค้า']
customer_name = Sales[Sales['ลูกค้า'] != 'None']['ลูกค้า'].iloc[0]
customer_Addre = Sales[Sales['ที่อยู่'] != 'None']['ที่อยู่'].iloc[0]
customer_phone = str(int(Sales[Sales['เบอร์โทร'] != 'None']['เบอร์โทร'].iloc[0]))
# customer_phone =customer_phone.astype(str)
########################
col1, col2 = st.columns([2, 1])
with col2:
    st.write('เลขที่ใบขายสินค้า: ',Sales_No)
    st.write(f"วันที่: {formatted_date}")
with col1:
    st.write("ลูกค้า:",customer_name)
    st.write("ที่อยู่:",customer_Addre)
    st.write("เบอร์โทร:",customer_phone )
st.write('---')
################################################
Sales=Sales[['รายการสินค้า',"จำนวนสินค้า",'ราคาขาย','ราคาขายรวม']]
# DCProd= DCProd.apply(pd.to_numeric, errors='coerce')
Sales[['ราคาขาย', 'ราคาขายรวม']] = Sales[['ราคาขาย', 'ราคาขายรวม']].apply(pd.to_numeric, errors='coerce')
Sales[['ราคาขาย', 'ราคาขายรวม']] = Sales[['ราคาขาย', 'ราคาขายรวม']].astype(float)
# Sales[['ราคาขายทุนเก่า', 'ราคาขายรวม']] = Sales[['ราคาขายทุนเก่า', 'ราคาขายรวม']].round(2)
# st.dataframe(Sales[['ราคาขายทุนเก่า', 'ราคาขายรวม']].style.format(precision=2))
filtered_sales=Sales
# st.table(Sales)
#################
รายการสินค้า = Sales['รายการสินค้า'].tolist()

# Filter the DataFrame to only include items in 'รายการสินค้า'
filtered_sales = Sales[Sales['รายการสินค้า'].isin(รายการสินค้า)]

# Reindex the DataFrame to match the order of 'รายการสินค้า' and reset the index
filtered_sales = Sales.set_index('รายการสินค้า').reindex(รายการสินค้า).reset_index()

# Add a sequential index starting from 1
filtered_sales.index = range(1, len(filtered_sales) + 1)
#######################################

# Convert to numeric first, handle errors by coercing to NaN
filtered_sales[['ราคาขาย', 'ราคาขายรวม']] = filtered_sales[['ราคาขาย', 'ราคาขายรวม']].apply(pd.to_numeric, errors='coerce')

# Calculate the sum of the 'ราคาขายรวม' column before formatting
Total = filtered_sales['ราคาขายรวม'].sum()

# Now format the 'ราคาขายทุนเก่า' and 'ราคาขายรวม' columns as strings with commas and two decimal places
filtered_sales[['ราคาขาย', 'ราคาขายรวม']] = filtered_sales[['ราคาขาย', 'ราคาขายรวม']].applymap(lambda x: f"{x:,.2f}")

# If needed, format the total as well
Total_formatted = f"{Total:,.2f}"
############################################

# Output the formatted total
filtered_sales=filtered_sales[['รายการสินค้า','จำนวนสินค้า','ราคาขาย','ราคาขายรวม']]
st.table(filtered_sales)
################
# Example values
label = 'รวมราคา:'
total = Total_formatted 

# Format the content with two decimal places
content = f"{label} {total}"

# Custom HTML with inline CSS
html_code = f"""
<div style="text-align: right; font-weight: bolder; background-color: #f0f0f0; padding: 10px; border-radius: 5px;">
    {content}
</div>
"""
import streamlit.components.v1 as components
components.html(html_code, height=100)
#########################
col3, col4= st.columns([3, 4])
with col3:
    ################
    # Example values
    label1 = 'ลูกค้า อนุมัติ/ได้รับสินค้า:______________________'

    # Format the content with two decimal places
    content = f"{label1}"

    # Custom HTML with inline CSS
    html_code = f"""
    <div style="text-align: left;background-color: #f0f0f0; padding: 20px;">
        {content}
    </div>
    """
    ############################
    label2 = 'ผู้จัดการร้านค้า อนุมัติ:_______________________ '
    import streamlit.components.v1 as components
    components.html(html_code, height=100)
    # Format the content with two decimal places
    content2 = f"{label2}"

    # Custom HTML with inline CSS
    html_code2 = f"""
    <div style="text-align: left;background-color: #f0f0f0; padding: 20px">
        {content2}
    </div>
    """
    import streamlit.components.v1 as components
    components.html(html_code2, height=100)
###############################
label4 = 'หมายเหตุ:___________________________________________________________________ '

# Format the content with two decimal places
content4 = f"{label4}"

# Custom HTML with inline CSS
html_code4 = f"""
<div style="text-align: center;background-color: #f0f0f0; padding: 50px">
    {content4}
</div>
"""
import streamlit.components.v1 as components
components.html(html_code4, height=100)
###############################
label3='ทางร้านขอกราบขอบพระคุณลูกค้าทุกท่าน ที่ให้ความกรุณาใช้สินค้าจากทางร้าน หวังว่าจะได้มีโอกาสรรับใช้ท่านอีกในโอกาสต่อไป ท่านสามารถติดต่อสอบถามสินค้าเพิ่มเติม หรือให้คำแนะนำทางร้านได้ที่ โทร 085 287 1408'

content3 = f"{label3}"

# Custom HTML with inline CSS
html_code3 = f"""
<div style="text-align: center; color: rgb(67, 66, 77); font-family: 'Times New Roman', Times, serif; font-size: 14px; background-color: #f0f0f0; padding: 10px; border-radius: 5px;">
    {content3}
</div>
"""
import streamlit.components.v1 as components
components.html(html_code3)



