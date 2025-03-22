import pandas as pd
import streamlit as st
import os
import numpy as np
from datetime import datetime
import io


st.title('MIB Product List Excel')
uploaded_file = st.file_uploader('Upload XLSX here:', type = ['xlsx'])


if uploaded_file:
    productlist_df_dict = pd.read_excel(uploaded_file, engine = 'openpyxl', sheet_name = None)

    colordf = productlist_df_dict['colors']
    productlist_df = productlist_df_dict['Current'].iloc[3:, :]



    productlist_df.columns = productlist_df.iloc[0]

    productlist_df = productlist_df[1:].reset_index(drop = True)

    Quickbooks_newproducts = pd.DataFrame(columns =  ['Product/service name', 'Category',	'Item type',	'SKU',	'Sales description',	'Sales price/rate', 'Income account',	'Purchase description',	'Purchase cost',	'Expense account',	'Quantity on hand',  'Quantity as-of date',	'Reorder point',	'Inventory asset account'])


    smallproductlist = productlist_df[productlist_df['NEW FOR SEASON'].notna()][['1X','2X','3X','4X','5X', '6X', '7X', '8X','NEW FOR SEASON', 'SKU', 'Description', 'Retail', 'Category', 'Country of Origin', 'hts code', 'Height', 'Width', 'Length', 'Weight']]

    sizes = ['1X','2X', '3X', '4X', '5X', '6X', '7X', '8X']
    newitems = []
    itemsku = []
    category = []
    purchasecosts = []
    sales = []
    totalnamelist = []
    countrylist = []
    heightlist = []
    widthlist = []
    lengthlist = []
    weightist = []
    htslist = []
    for index, row in smallproductlist.iterrows():
        sku = smallproductlist.loc[index, 'SKU']
        cat = 'Clothing:' +smallproductlist.loc[index, 'Category']
        sale = smallproductlist.loc[index, 'Retail']
        for j in sizes:
            purchasecost = smallproductlist.loc[index, j]
            if not pd.isna(smallproductlist.loc[index, j]):
                for k in [item.strip() for item in smallproductlist.loc[index, 'NEW FOR SEASON'].split(",")]: 

                    hts = smallproductlist.loc[index, 'hts code']
                    htslist.append(hts)
                    height= smallproductlist.loc[index, 'Height']
                    heightlist.append(height)
                    width = smallproductlist.loc[index, 'Width']
                    widthlist.append(width)
                    length = smallproductlist.loc[index, 'Length']
                    lengthlist.append(length)
                    weight = smallproductlist.loc[index, 'Weight']
                    weightist.append(weight)
                    newitem =smallproductlist.loc[index, 'Description'] +' '+ str(j)+  ' '+ str(k)
                    totalnames =smallproductlist.loc[index, 'Description']
                    country = smallproductlist.loc[index, 'Country of Origin']
                    totalnamelist.append(totalnames)
                    sku_2 = str(sku)+'-'+str(j)+'-'+str(k)
                    newitems.append(newitem)
                    itemsku.append(sku_2)
                    category.append(cat)
                    purchasecosts.append(purchasecost)
                    sales.append(sale)
                    countrylist.append(country)
    

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






    ####Second doctument 

    stringsl = 'stlye	desc    	color	color_desc	sizes	size_desc	coo	UPC	sku_ref		price	HEIGHT	WIDTH	LENGTH	WEIGHT'

    SKU_Upload_seconddoc = pd.DataFrame(columns = stringsl.split())

    def convert_df_to_csv(df):
    # Use StringIO to write to a string buffer (instead of a file)
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        return csv_buffer.getvalue()

    csv_data_qb = convert_df_to_csv(Quickbooks_newproducts)


    ####nbd start 
    colornamelist = []

    for color in [i.split('-')[2] for i in itemsku]:
        if colordf[['seasons color','Unnamed: 39']][colordf['seasons color'].str.lower() == color.lower()]['Unnamed: 39'].empty:
            colorname = 'missing'
        else:
            temp = colordf[['seasons color','Unnamed: 39']][colordf['seasons color'].str.lower() == color.lower()]['Unnamed: 39']
            
            colorname = temp.iloc[0]

        colornamelist.append(colorname)


    
    SKU_Upload_seconddoc['style']= [i.split('-')[0] for i in itemsku]
    SKU_Upload_seconddoc['desc'] = totalnamelist
    SKU_Upload_seconddoc['color'] = [i.split('-')[2] for i in itemsku]
    SKU_Upload_seconddoc['color_desc'] =  colornamelist
    SKU_Upload_seconddoc['sizes'] = [i.split('-')[1] for i in itemsku]
    SKU_Upload_seconddoc['size_desc']= [i.split('-')[1] for i in itemsku]
    SKU_Upload_seconddoc['coo'] = countrylist
    SKU_Upload_seconddoc['UPC']= itemsku
    SKU_Upload_seconddoc['sku_ref']= itemsku
    SKU_Upload_seconddoc['price'] =sales
    SKU_Upload_seconddoc['hts code']=  htslist
    SKU_Upload_seconddoc['HEIGHT'] =heightlist
    SKU_Upload_seconddoc['WIDTH'] =widthlist
    SKU_Upload_seconddoc['LENGTH'] =lengthlist
    SKU_Upload_seconddoc['WEIGHT'] =weightist


    SKU_Upload_seconddoc = SKU_Upload_seconddoc[['style', 'desc', 'color', 'color_desc', 'sizes', 'size_desc', 'coo',
       'UPC', 'sku_ref',  'hts code', 'price', 'HEIGHT', 'WIDTH', 'LENGTH', 'WEIGHT']]

    

    csv_data_seconddoc = convert_df_to_csv(SKU_Upload_seconddoc)

    #### Shopfiy start


    shopify = pd.DataFrame(columns = [
    "Handle", "Title", "Body (HTML)", "Vendor", "Product Category", "Type", "Tags", "Published",
    "Option1 Name", "Option1 Value", "Option1 Linked To", "Option2 Name", "Option2 Value", "Option2 Linked To",
    "Option3 Name", "Option3 Value", "Option3 Linked To", "Variant SKU", "Variant Grams",
    "Variant Inventory Tracker", "Variant Inventory Policy", "Variant Fulfillment Service",
    "Variant Price", "Variant Compare At Price", "Variant Requires Shipping", "Variant Taxable",
    "Variant Barcode", "Image Src", "Image Position", "Image Alt Text", "Gift Card",
    "SEO Title", "SEO Description", "Google Shopping / Google Product Category",
    "Google Shopping / Gender", "Google Shopping / Age Group", "Google Shopping / MPN",
    "Google Shopping / Condition", "Google Shopping / Custom Product",
    "Google Shopping / Custom Label 0", "Google Shopping / Custom Label 1",
    "Google Shopping / Custom Label 2", "Google Shopping / Custom Label 3",
    "Google Shopping / Custom Label 4", "Variant Image", "Variant Weight Unit",
    "Variant Tax Code", "Cost per item", "Included / United States", "Price / United States",
    "Compare At Price / United States", "Included / Canada", "Price / Canada",
    "Compare At Price / Canada", "Included / International", "Price / International",
    "Compare At Price / International", "Status"
    ])


    shopify['Title'] = totalnamelist
    shopify['Option1 Value'] = [i.split('-')[1] for i in itemsku]
    shopify['Option2 Value'] = [i.split('-')[2] for i in itemsku]
    shopify['Variant Price'] = sales
    shopify['Variant SKU'] = itemsku
    shopify['Variant Barcode'] = itemsku
    shopify['Cost per item'] = purchasecosts


    shopify['Variant Grams'] = 0.181436948
    shopify['Variant Inventory Tracker'] = 'shopify'
    shopify['Variant Inventory Policy'] = 'deny'
    shopify['Variant Fulfillment Service'] = 'nbd-vacaville'
    shopify['Variant Requires Shipping'] = 'TRUE'
    shopify['Variant Taxable'] = 'TRUE'
    shopify['Variant Weight Unit'] = 'lb'
    shopify['Handle']= shopify['Title'].str.replace(" ", "-").str.lower()

    shopify['Vendor'] = 'MIB Clothing'
    shopify['Product Category'] = 'Apparel & Accessories > Clothing'
    shopify['Type'] = [i.split(':',4)[1] for i in category]
    shopify['Option1 Name'] = 'Size'
    shopify['Option2 Name'] = 'Color'
    shopify['Gift Card'] = 'FALSE'
    shopify['Included / United States'] = 'TRUE'
    shopify['Included / Canada'] = 'TRUE'
    shopify['Included / International']= 'TRUE'
    shopify['Status']= 'draft'

    shopify.loc[shopify.duplicated(subset=['Handle']), ['Vendor', 'Product Category','Type',
                                                        'Option1 Name', 'Option2 Name',
                                                    'Gift Card', 'Included / United States',
                                                        'Included / Canada', 'Included / International',
                                                    'Status']] = ""



    csv_data_shopify = convert_df_to_csv(shopify)
    st.download_button(label ='Download Quickbooks CSV', data = csv_data_qb, file_name = 'Quickbooks_NewProductImport.csv', mime ='text/csv' )
    st.download_button(label ='Download NBD CSV', data = csv_data_seconddoc, file_name = 'NBDImport.csv', mime ='text/csv' )
    st.download_button(label ='Download Shopify CSV', data = csv_data_shopify, file_name = 'ShopifyImports.csv', mime ='text/csv' )

        




