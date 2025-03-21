import pandas as pd
import streamlit as st
import os
import numpy as np
from datetime import datetime


st.title('MIB Product List Excel')
uploaded_file = st.file_uploader('Upload XLSX here:', type = ['xlsx'])


if uploaded_file:
    df_dict = pd.read_excel(uploaded_file, engine = 'openpyxl', sheet_name = None)


    productlist_df = productlist_df_dict['Current'].iloc[3:, :]


    productlist_df.columns = productlist_df.iloc[0]

    productlist_df = productlist_df[1:].reset_index(drop = True)

    Quickbooks_newproducts = pd.DataFrame(columns =  ['Product/service name', 'Category',	'Item type',	'SKU',	'Sales description',	'Sales price/rate', 'Income account',	'Purchase description',	'Purchase cost',	'Expense account',	'Quantity on hand',  'Quantity as-of date',	'Reorder point',	'Inventory asset account'])


    smallproductlist = productlist_df[productlist_df['NEW FOR SEASON'].notna()][['2X','3X','4X','5X', '6X', '7X', '8X','NEW FOR SEASON', 'SKU', 'Description', 'Retail', 'Category']]

    sizes = ['2X', '3X', '4X', '5X', '6X', '7X', '8X']
    newitems = []
    itemsku = []
    category = []
    purchasecosts = []
    sales = []
    for index, row in smallproductlist.iterrows():
        sku = smallproductlist.loc[index, 'SKU']
        cat = 'Clothing:' +smallproductlist.loc[index, 'Category']
        sale = smallproductlist.loc[index, 'Retail']
        for j in sizes:
            purchasecost = smallproductlist.loc[index, j]
            if not pd.isna(smallproductlist.loc[index, j]):
                for k in [item.strip() for item in smallproductlist.loc[index, 'NEW FOR SEASON'].split(",")]: 
                    newitem =smallproductlist.loc[index, 'Description'] +' '+ str(j)+  ' '+ str(k)
                    sku_2 = str(sku)+'-'+str(j)+'-'+str(k)
                    newitems.append(newitem)
                    itemsku.append(sku_2)
                    category.append(cat)
                    purchasecosts.append(purchasecost)
                    sales.append(sale)
    

    Quickbooks_newproducts['Product/service name'] = newitems
    Quickbooks_newproducts['SKU'] = itemsku
    Quickbooks_newproducts['Category'] = category
    Quickbooks_newproducts['Purchase cost'] = purchasecosts
    Quickbooks_newproducts['Sales price/rate'] = sales


    Quickbooks_newproducts['Item type'] = 'Inventory'
    #Quickbooks_newproducts['Purchase description'] = 'INCOME'
    Quickbooks_newproducts['Income account'] = '4010 Sales:Merchandise'
    Quickbooks_newproducts['Expense account'] = '5005 Cost of Good Sold:COGS - Finished Goods'
    Quickbooks_newproducts['Quantity on hand'] = 0
    #Quickbooks_newproducts['Reorder point'] = np.nan
    Quickbooks_newproducts['Inventory asset account'] = '1450 Inventory:Finished Goods'
    Quickbooks_newproducts['Quantity as-of date'] = datetime.today().strftime("%m/%d/%Y")

    def convert_df_to_csv(df):
    # Use StringIO to write to a string buffer (instead of a file)
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        return csv_buffer.getvalue()

    csv_data = convert_df_to_csv(final_df)
        
    st.download_button(label ='Download CSV', data = csv_data, file_name = 'Quickbooks_NewProductImport.csv', mime ='text/csv' )

        




