import streamlit as st
import pandas as pd
import yfinance as yf
import plotly.express as px

#excel_path = 'C:/Users/user/Desktop/MyScripts/Index Bubble Chart/Index Weight.xlsx'
# Dropbox direct download link
excel_path = 'https://dl.dropboxusercontent.com/s/j6jjjggzxs3agwy1r6egd/Index-Weight.xlsx?dl=1'

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
    # Refresh button in the sidebar
    if st.sidebar.button('Refresh Data'):
        st.experimental_rerun()

    data = load_data(excel_path)
    processed_data = {name: fetch_and_calculate(df) for name, df in data.items()}

    for name, df in processed_data.items():
        df['Today Pct Change'] = df['Today Pct Change'].apply(format_pct_change)

    index_choice = st.sidebar.selectbox('Select Index', list(processed_data.keys()))
    df_display = processed_data[index_choice]

    def color_scale(val):
        if val > 5: return 'red'
        elif val > 4: return 'Crimson'
        elif val > 3: return 'pink'
        elif val > 2: return 'brown'
        elif val > 1: return 'orange'
        else: return 'blue'

    df_display['Color'] = df_display['Volume Ratio'].apply(color_scale)

    # Set the y-axis to include negative returns
    min_pct_change = df_display['Today Pct Change'].str.rstrip('%').astype('float').min()
    max_pct_change = df_display['Today Pct Change'].str.rstrip('%').astype('float').max() * 1.1  # 10% padding

    fig = px.scatter(df_display, x='Volume Ratio', y='Today Pct Change', size='Weight',
                     hover_data=['Name', 'Code', 'Today Pct Change', 'Volume Ratio'],
                     color='Color', color_discrete_map="identity",
                     title=f'Bubble Chart for {index_choice}', height=800, width=1000)

    # Apply a logarithmic scale to the x-axis if all values are positive
    if df_display['Volume Ratio'].min() > 0:
        fig.update_xaxes(type='log')
    
    # Adjust the range of the y-axis to include negative returns
    fig.update_xaxes(
        zeroline=True, zerolinewidth=2, zerolinecolor='black',
        tickfont=dict(size=16, color='black'),  # Update tick font size here
        tickformat='0.2f'  # Formats the tick labels to float with leading zero
    )

    fig.update_yaxes(
        zeroline=True, zerolinewidth=2, zerolinecolor='black',
        tickfont=dict(size=16, color='black')  # Update tick font size here
    )

    # Update layout for axis titles
    fig.update_layout(
        xaxis_title="Volume Ratio",
        yaxis_title="Today Pct Change",
        font=dict(
            family="Courier New, monospace",
            size=18,
            color="RebeccaPurple"
        ),
        xaxis=dict(
            title_font=dict(size=18, color='RebeccaPurple'),
            tickfont=dict(size=16, color='RebeccaPurple'),  # Update tick font size here
            zeroline=True, zerolinewidth=2, zerolinecolor='black'
        ),
        yaxis=dict(
            title_font=dict(size=18, color='RebeccaPurple'),
            tickfont=dict(size=16, color='RebeccaPurple'),  # Update tick font size here
            zeroline=True, zerolinewidth=2, zerolinecolor='black'
        )
    )

    st.plotly_chart(fig)

if __name__ == "__main__":
    main()




