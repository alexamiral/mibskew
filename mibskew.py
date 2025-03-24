import pandas as pd
import streamlit as st
import io
import openpyxl
import plotly.graph_objects as go



st.title('MIB Skew Orders')
uploaded_file = st.file_uploader('Upload XLSX here:', type = ['xlsx'])

if uploaded_file:
    df_dict = pd.read_excel(uploaded_file, engine = 'openpyxl', sheet_name = None)
    
    for n in df_dict:
        globals()[f'df_{n}'] = df_dict[n]

    for i in df_dict: 
        header_row_index = globals()[f'df_{i}'][globals()[f'df_{i}'].eq("Color").any(axis=1)].index[0] 
        
        final_row_index = globals()[f'df_{i}'][globals()[f'df_{i}'].eq( 'ALL SIZES MUST BE TO "MAKING IT BIG SPECS"').any(axis=1)].index[0]
        num_rows  = final_row_index - header_row_index
        globals()[f'table_{i}'] = globals()[f'df_{i}'].iloc[header_row_index: header_row_index+num_rows].reset_index(drop = True)
        globals()[f'table_{i}'].columns = globals()[f'table_{i}'].iloc[0]
        globals()[f'table_{i}'] = globals()[f'table_{i}'][1:].reset_index(drop = True)
        globals()[f'table_{i}'] = globals()[f'table_{i}'].iloc[:, globals()[f'table_{i}'].columns.get_loc('Color'):globals()[f'table_{i}'].columns.get_loc('Total')]
        globals()[f'table_{i}'] = globals()[f'table_{i}'].dropna(axis=1, how='all')



    newcodelist = []
    for i in df_dict:
        for rindex, row in globals()[f'table_{i}'].iterrows():
            for cname, value in row.items():
                if isinstance(value, int) and value>0:
                    #print(value, rindex, cname)
                    temp = globals()[f'table_{i}'].loc[rindex, 'For SKU']
                    newcode = f"{i}-{cname}-{temp}"
                    for m in range(0, value):
                        newcodelist.append(newcode)

    final_df = pd.DataFrame({'Orders:': newcodelist})
    def convert_df_to_csv(df):
    # Use StringIO to write to a string buffer (instead of a file)
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        return csv_buffer.getvalue()

    csv_data = convert_df_to_csv(final_df)



    ##### Po downloads 
    for i in df_dict:
        globals()[f'po_df_{i}'] = pd.DataFrame(columns = 'PURCHASE_ORDER_NUMBER	LINE_NUMBER	PRODUCT	QUANTITY	UNIT_OF_MEASURE_CODE	UNIT_OF_MEASURE_QTY	HOST_LINE_NUMBER'.split()
        )

    ponewlines = {}
    counter = 1
    for i in df_dict:
        for rindex, row in globals()[f'table_{i}'].iterrows():
                for cname, value in row.items():
                    if isinstance(value, int):
                        temp = globals()[f'table_{i}'].loc[rindex, 'For SKU']
                        ponewlines[counter] = {}
                        ponewlines[counter]['PURCHASE_ORDER_NUMBER']=globals()[f'df_{i}'].iloc[0,:]['Unnamed: 13']
                        ponewlines[counter]['LINE_NUMBER']=counter
                        ponewlines[counter]['PRODUCT']=f"{str(i)}-{cname}-{temp}"
                        ponewlines[counter]['QUANTITY']=value
                        ponewlines[counter]['UNIT_OF_MEASURE_CODE']='EACH'
                        ponewlines[counter]['UNIT_OF_MEASURE_QTY']=1
                        ponewlines[counter]['HOST_LINE_NUMBER']=counter
                        counter +=1

    podownload = pd.DataFrame.from_dict(ponewlines, orient = 'index')
    csvpodownload = convert_df_to_csv(podownload)

    def isint(x):
        if isinstance(x, int) and x>0:
            return x
    


    itemcountdict = {}
    for i in df_dict:
        itemcount = 0
        for index, row in globals()[f'table_{i}'].iterrows():
            itemcount += row.map(isint).sum()
        itemcountdict[i] = itemcount
        globals()[f'itemcount_{i}'] = itemcount


    df_dict_keys = list(df_dict.keys())
    df_dict_keys.append('TOTAL')

    itemcountvalues =list(itemcountdict.values())
    itemcountvalues.append(sum(itemcountvalues))


    fig = go.Figure(data=[go.Table(
    header=dict(values=['SKU', 'Item Count'],
    fill_color='blue'),
                 cells=dict(values=[df_dict_keys,itemcountvalues ], fill_color ='lightblue'
                           ))
                     ])
    




    st.write(fig)
    st.download_button(label ='MIB SKU CSV', data = csv_data, file_name = 'MIB_Order_SKUs.csv', mime ='text/csv' )
    st.download_button(label ='PO Download CSV', data = csvpodownload, file_name = 'Po_Downloads.csv', mime ='text/csv' )

