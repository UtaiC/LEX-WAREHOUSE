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

# Function to load the CSS file #############################################
def load_css(file_name):
    with open(file_name) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Function to load custom HTML content######################################
def load_html(html_code):
    components.html(html_code)

# Load the CSS file #########################################################
load_css("page.css")
###############################
st.sidebar.subheader('Main Manu')
page = st.sidebar.radio("Go to", ["Home", "Stock Update", "Sales Order","Print Order"])
##########################

Banner=Image.open('WATTANA-Banner.jpg')
st.image(Banner)

## Files Read ###############################################
file='https://docs.google.com/spreadsheets/d/10XwWMOqFxrQMiz_d7Vv__806167AQ8Wy6v82z9MOsQc/export?format=xlsx'
stock_data = pd.read_excel(file, header=0, engine='openpyxl')
stock_data.index=stock_data.index+1
stock_data=stock_data.dropna()
############################
Sales_Files='https://docs.google.com/spreadsheets/d/1tNRSxI9xv0FJMHbmwLpszP-ObAgpb0dY5AIciPmZWWY/export?format=xlsx'
Sales = pd.read_excel(Sales_Files, header=0, engine='openpyxl')
Sales['รายการสินค้า']=Sales['รายการสินค้า'].astype(str)
Sales['รายการสินค้า']=Sales['รายการสินค้า'].str.split('|').str[0]
##################################################
# Clean the 'รายการสินค้า' column in Sales
Sales["รายการสินค้า"] = Sales["รายการสินค้า"].str.strip()
Sales["รายการสินค้า"] = Sales["รายการสินค้า"].str.replace(r'\s+', ' ', regex=True)
Sales["รายการสินค้า"] = Sales["รายการสินค้า"].str.lower()
# Clean the 'รายการ' column in stock_data
stock_data["รายการ"] = stock_data["รายการ"].str.strip()
stock_data["รายการ"] = stock_data["รายการ"].str.replace(r'\s+', ' ', regex=True)
stock_data["รายการ"] = stock_data["รายการ"].str.lower()
########################################################
# Perform the merge
Sales = pd.merge(Sales, stock_data, left_on='รายการสินค้า', right_on='รายการ', how='left')
Sales.index = Sales.index + 1
# Home page #################################################
if page == "Home":
    st.subheader("Welcome to the LEX WAREHOUSE")
    Lex01=Image.open('LEX-Pics01.jpg')
    st.image(Lex01,width=695)

# Stock Page ################################################
elif page == "Stock Update":
    st.subheader("รายการสินค้า")

    stock_data
    ############# Seek List ###########################


    # Streamlit input box for keywords
    keywords_input = st.text_input("กรุณากรอกรายการสินค้าที่ต้องการค้นหา (ใช้เครื่องหมาย ',' คั้นคำที่ต้องการค้นหา หากต้องการค้นหาหลายรายการ):")

    # Function to filter DataFrame by keywords
    def filter_by_keywords(stock_data, keywords):
        pattern = '|'.join(keywords)
        filtered_stock_data = stock_data[stock_data['รายการ'].str.contains(pattern, case=False, na=False)]
        return filtered_stock_data

    # Check if user has input any keywords
    if keywords_input:
        # Convert user input into a list of keywords
        keywords = [kw.strip() for kw in keywords_input.split(',')]
        # Filter the DataFrame
        filtered_stock_data = filter_by_keywords(stock_data, keywords)
        # Display the filtered DataFrame
        st.write("รายการสินค้าที่ท่านค้นหา:")
        st.write(filtered_stock_data)
    else:
        st.write("กรุณากรอกรายกสินค้าที่ต้องการค้นหา")

# Sales Page #################################################
elif page == "Sales Order":
    st.subheader("ตรวจสอบรายการขาย")
    ############################
    st.write('กรุณาเลือก Sales_Date และ Sales_No เพื่อตรวจสอบรายการขายให้ถูกต้องก่อนดำเนินการขั้นตอนต่อไป')
    Sales_Date = st.selectbox("Timestamp",['2024-01',	'2024-02',	'2024-03',	'2024-04',	'2024-05',	'2024-06',	'2024-07',	'2024-08',	'2024-09',	'2024-10',	'2024-11',	'2024-12'] )
    #############################
    
    Sales_No = st.selectbox('Sales_No',['WH-000','WH-001', 'WH-002', 'WH-003', 'WH-004', 'WH-005', 'WH-006', 'WH-007', 'WH-008', 'WH-009', 'WH-010', 'WH-011', 'WH-012', 'WH-013', 'WH-014', 'WH-015', 'WH-016', 'WH-017', 'WH-018', 'WH-019', 'WH-020', 'WH-021', 'WH-022', 'WH-023', 'WH-024', 'WH-025', 'WH-026', 'WH-027', 'WH-028', 'WH-029', 'WH-030', 'WH-031', 'WH-032', 'WH-033', 'WH-034', 'WH-035', 'WH-036', 'WH-037', 'WH-038', 'WH-039', 'WH-040', 'WH-041', 'WH-042', 'WH-043', 'WH-044', 'WH-045', 'WH-046', 'WH-047', 'WH-048', 'WH-049', 'WH-050'] )
    
    #########################################################
    Sales["เลขที่เสนอราคา"]=Sales["เลขที่เสนอราคา"].fillna('None')
    Sales["Timestamp"]=Sales["Timestamp"].astype(str)
    Sales=Sales[Sales["Timestamp"].str.contains(Sales_Date)]
    Sales=Sales[Sales["เลขที่เสนอราคา"].str.contains(Sales_No)]
    ##############################
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
    ################################################################

    # Convert 'Timestamp' to just date (year, month, day)
    Sales['Date'] = pd.to_datetime(Sales['Timestamp']).dt.date

    # Step 1: Identify rows where 'ประเภทลูกค้า' is 'ลูกค้าช่าง'
    condition = Sales['ประเภทลูกค้า'] == 'ลูกค้าช่าง'

    # Step 2: Find unique 'เลขที่เสนอราคา' and 'Date' for these rows
    duplicate_rows = Sales[condition]
    unique_criteria = duplicate_rows[['เลขที่เสนอราคา', 'Date']].drop_duplicates()

    # Step 3: For each unique (เลขที่เสนอราคา, Date), update 'ประเภทลูกค้า' to 'ลูกค้าช่าง'
    for _, row in unique_criteria.iterrows():
        Sales.loc[
            (Sales['เลขที่เสนอราคา'] == row['เลขที่เสนอราคา']) & 
            (Sales['Date'] == row['Date']), 
            'ประเภทลูกค้า'
        ] = 'ลูกค้าช่าง'
    Sales=Sales[['Timestamp','จำนวนสินค้า','เลขที่เสนอราคา','รหัส','รายการสินค้า','รายการ','ประเภทลูกค้า','จำนวนเก่า','ราคาทุนเก่า','ราคาขายทุนเก่า','ราคาช่างทุนเก่า','จำนวนใหม่','ราคาทุนใหม่','ราคาขายทุนใหม่','ราคาช่างทุนใหม่','ยอดรวมสต๊อก']]
    #########################################################
    # Convert columns to numeric values and handle NaN
    Sales['จำนวนเก่า'] = Sales['จำนวนเก่า'].apply(pd.to_numeric, errors='coerce')
    Sales['จำนวนใหม่'] = Sales['จำนวนใหม่'].apply(pd.to_numeric, errors='coerce')
    Sales['ประเภทลูกค้า'] = Sales['ประเภทลูกค้า'].fillna('None')

    # Calculate AMT based on customer type
 

    # Condition for customer type 'ลูกค้าช่าง'
    is_customer_technician = Sales['ประเภทลูกค้า'] == 'ลูกค้าช่าง'

    # Condition for old stock being less than the required quantity
    is_old_stock_less = Sales['จำนวนเก่า'] < Sales["จำนวนสินค้า"]

    # Calculate ราคาขายรวม based on conditions
    Sales['ราคาขายรวม'] = np.where(
        is_customer_technician & is_old_stock_less,
        Sales['ราคาช่างทุนเก่า'] * Sales['จำนวนเก่า'] + 
        Sales['ราคาช่างทุนใหม่'] * (Sales["จำนวนสินค้า"] - Sales['จำนวนเก่า']),
        np.where(is_customer_technician,Sales['ราคาช่างทุนเก่า'] * Sales["จำนวนสินค้า"],
        np.where(is_old_stock_less,Sales['ราคาขายทุนเก่า'] * Sales['จำนวนเก่า'] + Sales['ราคาขายทุนใหม่'] * (Sales["จำนวนสินค้า"] - Sales['จำนวนเก่า']),
                 Sales['ราคาขายทุนเก่า'] * Sales["จำนวนสินค้า"])))

    # Calculate unit price
    Sales['ราคาขาย'] = Sales['ราคาขายรวม'] / Sales["จำนวนสินค้า"]

    # Display the result
    Sales

    ### Revise Sales Unit #############################################

    # Identify overstocked items
    overstocked = Sales.merge(stock_data, on='รายการ')
    overstocked = overstocked[overstocked['จำนวนสินค้า'] > overstocked['ยอดรวมสต๊อก']]

    # Create a dictionary to store the new quantities
    new_quantities = {}

    # Display and update overstocked items
    if not overstocked.empty:
        st.write("**แก้ไขจำนวนสินค้า ตามจำนวนสินค้าคงเหลือในสต๊อก:**")
        
        for index, row in overstocked.iterrows():
            # Ensure all values are of type int (or float)
            max_stock = int(row['ยอดรวมสต๊อก'])
            current_quantity = int(row['จำนวนสินค้า'])
            
            # Adjust current_quantity if it exceeds max_stock
            if current_quantity > max_stock:
                # st.warning(f"The current quantity for {row['รายการ']} exceeds the available stock. Setting to max stock.")
                current_quantity = max_stock

            new_quantity = st.number_input(
                f"ปรับแก้ไขจำนวนที่สามารถขายได้ สำหรับสินค้ารายการ:  {row['รายการ']} (current: {current_quantity}, max stock: {max_stock})",
                min_value=0,
                max_value=max_stock,
                value=current_quantity,
                step=1
            )
            # Store the new quantity in the dictionary
            new_quantities[row['รายการ']] = new_quantity

        # Update the Sales DataFrame with the new quantities using replace
        if st.button("แก้ไขรายการ"):
            for item, new_qty in new_quantities.items():
                Sales['จำนวนสินค้า'] = Sales['จำนวนสินค้า'].mask(Sales['รายการ'] == item, new_qty)
            
            st.success("Sales quantities updated successfully!")
            st.write("Updated Sales DataFrame:")
            st.write(Sales[['เลขที่เสนอราคา','ลูกค้า','ที่อยู่','เบอร์โทร','จำนวนสินค้า','รายการสินค้า']])
        ##### Profit #####################################################
    Total_Sales=Sales['ราคาขายรวม'].sum()
    Total_Sales=Total_Sales.round(2)
    st.write('รวมราคาขาย:',Total_Sales,'บาท')
    #############################################################################
# Calculate Cost_Old based on the same conditions as ราคาขายรวม
    Cost_Old = np.where(
        is_customer_technician & is_old_stock_less,
        (Sales['ราคาช่างทุนเก่า'] * (Sales['จำนวนเก่า']-Sales['จำนวนใหม่']).sum()),
        np.where(is_customer_technician,
                (Sales['ราคาช่างทุนเก่า'] * Sales["จำนวนสินค้า"]).sum(),
                np.where(is_old_stock_less,
                        (Sales['ราคาทุนเก่า'] * (Sales['จำนวนเก่า']-Sales['จำนวนใหม่']).sum()).sum(),
                        (Sales['ราคาทุนเก่า'] * Sales["จำนวนสินค้า"]).sum()))
    )

    # Calculate Cost_New only where old stock is less than required quantity
    Cost_New = np.where(
        is_customer_technician & is_old_stock_less,
        (Sales['ราคาช่างทุนเก่า'] * (Sales['จำนวนเก่า']-Sales['จำนวนใหม่']).sum()),
        np.where(is_customer_technician,
                (Sales['ราคาช่างทุนเก่า'] * Sales["จำนวนสินค้า"]).sum(),
                np.where(is_old_stock_less,
                        (Sales['ราคาทุนเก่า'] * (Sales['จำนวนเก่า']-Sales['จำนวนใหม่']).sum()).sum(),
                        (Sales['ราคาทุนเก่า'] * Sales["จำนวนสินค้า"]).sum()))
    )
    #############################################################################
    Cost_Old=Cost_Old.sum()
    Cost_Old=Cost_Old.round(2)
    st.write('รวมต้นทุนเก่า:',Cost_Old,'บาท')
    Cost_New=Cost_New.sum()
    Cost_New=Cost_New.round(2)
    st.write('รวมต้นทุนใหม่:',Cost_New,'บาท')
    TT_Cost=Cost_Old+Cost_New
    st.write('รวมต้นทุนทั้งสิ้น:',TT_Cost,'บาท')
    Magine=Total_Sales-TT_Cost
    Magine=Magine.round(2)
    st.write('กำไรขั้นต้น:',Magine,'บาท')
    PCT_Magine=100-((TT_Cost/Total_Sales)*100)
    PCT_Magine=PCT_Magine.round(2)
    st.write('กำไรขั้นต้น-%:',PCT_Magine,'%')
    if Sales.empty:  # Check if the DataFrame is empty
        st.write("")
    else:

        Sales.to_excel(Sales_No+'Sales-Update.xlsx')
        st.write('Sales-updated and Exportsuccessfully!')
    if Sales_No == 'WH-000':
        st.write("ท่านยังไม่ได้เลือก Sales_No ที่ถูกต้องเพื่อตรวจสอบรายการขาย")
    elif Sales.empty:  # Check if the DataFrame is empty
        st.write("โปรดเลือกรายการขายสินค้าให้ถูกต้อง หรือไม่ได้บันทึกข้อมูลการขาย")
    else:
        st.write("กรุณาดำเนินตามกระบวนการขายต่อ")
###########################################################################################
elif page == "Print Order":
    ##############
    Logo=Image.open('WATTANA-Logo-Sales.jpg')
    st.image(Logo,width=700)
    #########################

    #########################################################
    Sales["เลขที่เสนอราคา"]=Sales["เลขที่เสนอราคา"].fillna('None')
    Sales["Timestamp"]=Sales["Timestamp"].astype(str)
    Sales=Sales[Sales["Timestamp"].str.contains(Sales_Date)]
    Sales=Sales[Sales["เลขที่เสนอราคา"].str.contains(Sales_No)]
    ##############################
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
    Sales["Timestamp"] = pd.to_datetime(Sales["Timestamp"], errors='coerce')
    Sales = Sales.dropna(subset=["Timestamp"])
    Sales = Sales[Sales["Timestamp"].dt.strftime('%Y-%m-%d').str.startswith(Sales_Date)]
    Sales=Sales[Sales["เลขที่เสนอราคา"].str.contains(Sales_No)]

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



