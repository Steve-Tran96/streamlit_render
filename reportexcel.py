#
# Installing streamlit, openpyxl, pandas, numpy, plotly-express

import pandas as pd
import streamlit as st
import plotly.express as px

import streamlit as st
from streamlit_autorefresh import st_autorefresh

from datetime import date


st.set_page_config(page_title="Sales Dashboard", page_icon=":bar_chart:", layout="wide")
st.title('Quản lý kho hàng')

# update every 5 mins
st_autorefresh(interval=0.5 * 60 * 1000, key="dataframerefresh")

excel_filename = 'QuanLyKho.xlsm'
sheetname = 'Kho'

def get_data():
    df = pd.read_excel(
        io='QuanLyKho.xlsm', # ten file
        sheet_name= sheetname, #sheet name cua file
        engine='openpyxl', #dung thu vien openpyxl
        skiprows=8, # bo qua 8 dong dau trong sheet
        usecols='B:L', # chon tu cot B den L
        nrows=1000,
    )    
    return df

def get_vt_data():
    df_vattu = pd.read_excel(
        io=excel_filename, 
        sheet_name= sheetname,
        engine='openpyxl',
        skiprows=8,
        usecols= [1,10], # chon tu cot B den K
        nrows=100,
    )    
    return df_vattu

pie_chart = px.pie(get_vt_data(),
            title = 'Số lượng tồn kho',
            values = 'SL Tồn',
            names = 'Mã sản phẩm'
)

if __name__ == '__main__':    
    col1, col2, col3= st.columns([2,1,2])
    with col1:
        st.markdown("<h1 style='margin-bottom: 9px;'>Kho vật tư của tháng " + str(date.today())+"</h1>", unsafe_allow_html=True)
        # st.subheader('Kho vật tư của tháng ' + str(date.today()))
        st.dataframe(get_data())
    with col2:
        st.subheader('Mã sản phẩm ')
        st.dataframe(get_vt_data()) 
    with col3:
        st.plotly_chart(pie_chart)

