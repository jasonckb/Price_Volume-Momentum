import streamlit as st
import pandas as pd
import yfinance as yf
import plotly.express as px
import openpyxl

#excel_path = 'C:/Users/user/Desktop/MyScripts/Index Bubble Chart/Index Weight.xlsx'
# Dropbox direct download link
excel_path = 'https://www.dropbox.com/scl/fi/nw5fpges55aff7x5q3gh9/Index-Weight.xlsx?rlkey=rxdopdklplz15jk97zu2sual5&dl=1'

@st.cache_data(show_spinner=False)
def load_data(excel_path):
    sheet_names = ['HSI', 'HSTECH', 'HSCEI']
    dtype = {'Code': str}
    return {name: pd.read_excel(excel_path, sheet_name=name, dtype=dtype) for name in sheet_names}

@st.cache_data(show_spinner=False)
def fetch_and_calculate(df):
    for index, row in df.iterrows():
        stock_code = f"{row['Code'].zfill(4)}.HK"
        stock = yf.Ticker(stock_code)
        hist = stock.history(period="11d")

        if hist.empty:
            df.at[index, 'Today Pct Change'] = None
            df.at[index, 'Volume Ratio'] = None
        else:
            today_data = hist.iloc[-1]
            avg_volume_10d = hist['Volume'][:-1].mean()
            df.at[index, 'Today Pct Change'] = round(((today_data['Close'] - today_data['Open']) / today_data['Open']) * 100, 2)
            df.at[index, 'Volume Ratio'] = round(today_data['Volume'] / avg_volume_10d, 2)
    
    return df

def format_pct_change(val):
    return f"{val}%" if pd.notnull(val) else ""

def main():
    st.set_page_config(page_title="HK Stocks Volume & Price Momentum by Jason Chan", layout="wide")
    st.title('Index Components Volume & Price Momentum by Jason Chan')

    # 1. Toggle for showing/hiding the bubble chart
    show_chart = st.sidebar.checkbox("Show Bubble Chart", value=True)

    # Refresh button in the sidebar
    if st.sidebar.button('Refresh Data'):
        st.experimental_rerun()

    # Fetch and process the data...
    if 'raw_data' not in st.session_state:
        st.session_state.raw_data = load_data(excel_path)
    
    processed_data = {name: fetch_and_calculate(st.session_state.raw_data[name].copy(deep=True)) 
                      for name in st.session_state.raw_data}

    index_choice = st.sidebar.selectbox('Select Index', list(processed_data.keys()))
    df_display = processed_data[index_choice].copy(deep=True)

    # 3. Slider for adjusting bubble size (demonstrative purpose)
    size_multiplier = st.sidebar.slider("Bubble Size Multiplier", min_value=0.5, max_value=5.0, value=1.0)

    if show_chart:
        # Check and prepare the data
        df_display['Volume Ratio'] = pd.to_numeric(df_display['Volume Ratio'], errors='coerce').fillna(0)
        df_display['Today Pct Change'] = pd.to_numeric(df_display['Today Pct Change'].str.rstrip('%'), errors='coerce').fillna(0)
        df_display['Size'] = df_display['Weight'] * size_multiplier  # Adjust 'Weight' by slider value
        
        # Ensure no NaN values for 'Size'
        df_display['Size'] = pd.to_numeric(df_display['Size'], errors='coerce').fillna(0)

        # Now creating the figure
        fig = px.scatter(
            df_display, 
            x='Volume Ratio', 
            y='Today Pct Change', 
            size='Size',
            hover_data=['Name', 'Code', 'Today Pct Change', 'Volume Ratio'],
            color='Color', 
            color_discrete_map="identity",
            title='Bubble Chart', 
            height=800, 
            width=1000
        )

        # Customization for a black background and white text...
        fig.update_layout({
            'plot_bgcolor': 'black',
            'paper_bgcolor': 'black',
            'font': {'color': 'white'}
        })
        

        fig.update_xaxes(
            zeroline=True, zerolinewidth=2, zerolinecolor='white',
            tickfont=dict(size=16, color='white'),
            title_font=dict(size=18, color='white')
        )

        fig.update_yaxes(
            zeroline=True, zerolinewidth=2, zerolinecolor='white',
            tickfont=dict(size=16, color='white'),
            title_font=dict(size=18, color='white')
        )

        st.plotly_chart(fig)

if __name__ == "__main__":
    main()




