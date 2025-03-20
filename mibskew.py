
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
    
    st.download_button(label ='Download CSV', data = csv_data, file_name = 'MIB_Order_SKUs.csv', mime ='text/csv' )