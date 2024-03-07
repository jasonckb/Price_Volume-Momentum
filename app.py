import streamlit as st
import pandas as pd
import yfinance as yf
import plotly.express as px
import plotly.graph_objs as go
import openpyxl
import datetime

# Excel data path
excel_path = 'https://www.dropbox.com/scl/fi/nw5fpges55aff7x5q3gh9/Index-Weight.xlsx?rlkey=rxdopdklplz15jk97zu2sual5&dl=1'

# Define a function to load Excel data, using cache for historical data persistence.
@st.cache(show_spinner=False)
def load_historical_data(excel_path):
    sheet_names = ['HSI', 'HSTECH', 'HSCEI', 'SP 500']
    dtype = {'Code': str}
    return {name: pd.read_excel(excel_path, sheet_name=name, dtype=dtype) for name in sheet_names}

# Define a function to fetch and calculate historical data, applying caching.
@st.cache(show_spinner=False)
def fetch_and_calculate_historical(df, index_name):
    return process_data(df, index_name, period='max')

# Define a function to fetch and calculate today's data, excluding it from caching.
def fetch_and_calculate_today(df, index_name):
    return process_data(df, index_name, period='1d')

# Process the fetched data accordingly.
def process_data(df, index_name, period):
    # Define the end date based on the period specified.
    end_date = 'today' if period == '1d' else None

    for index, row in df.iterrows():
        stock_code = row['Code'] if index_name == 'SP 500' else f"{row['Code'].zfill(4)}.HK"
        stock = yf.Ticker(stock_code)
        hist = stock.history(period=period, end=end_date)

        if not hist.empty and len(hist) > 1:
            today_close = hist.iloc[-1]['Close']
            prev_close = hist.iloc[-2]['Close']
            today_pct_change = round(((today_close - prev_close) / prev_close) * 100, 2)
            df.at[index, 'Today Pct Change'] = today_pct_change
            df.at[index, 'Volume Ratio'] = round(hist.iloc[-1]['Volume'] / hist['Volume'][:-1].mean(), 2)
    
    return df

def color_scale(val):
    if val > 5: return 'red'
    elif val > 4: return 'Crimson'
    elif val > 3: return 'pink'
    elif val > 2: return 'brown'
    elif val > 1: return 'orange'
    else: return 'gray'

def generate_plot(df_display, index_choice):
    fig = px.scatter(
        df_display, 
        x='Volume Ratio', 
        y='Today Pct Change', 
        size='Weight',
        hover_data=['Name', 'Code', 'Today Pct Change', 'Volume Ratio'],
        color='Color', 
        color_discrete_map="identity",
        title=f"{index_choice} Volume Ratio: Today VS.10 Days Average",
        height=800, 
        width=1000
    )

    fig.update_layout(
        plot_bgcolor='black',
        paper_bgcolor='black',
        font=dict(color='white'),
        xaxis=dict(title_font=dict(color='white'), tickfont=dict(color='white')),
        yaxis=dict(title_font=dict(color='white'), tickfont=dict(color='white'))
    )

    fig.update_xaxes(type='log' if df_display['Volume Ratio'].min() > 0 else 'linear')
    return fig

# Main function orchestrating the app flow.
def main():
    st.set_page_config(page_title="Index Constituents Volume & Price Momentum", layout="wide")
    st.title('Index Components Volume & Price Momentum')

    if 'raw_data' not in st.session_state:
        st.session_state['raw_data'] = load_historical_data(excel_path)
    
    index_choice = st.sidebar.selectbox('Select Index', ['HSI', 'HSTECH', 'HSCEI', 'SP 500'])

    if st.sidebar.button('Daily Historical Update'):
        st.session_state['processed_data'] = {
            name: fetch_and_calculate_historical(st.session_state['raw_data'][name].copy(), name) 
            for name in st.session_state['raw_data']
        }

    if st.sidebar.button('Intraday Refresh'):
        st.session_state['processed_data'][index_choice] = fetch_and_calculate_today(
            st.session_state['raw_data'][index_choice].copy(), index_choice)

    df_display = st.session_state['processed_data'].get(index_choice, pd.DataFrame()).copy()
    df_display['Color'] = df_display['Volume Ratio'].apply(color_scale)

    if not df_display.empty:
        fig = generate_plot(df_display, index_choice)
        st.plotly_chart(fig)

if __name__ == "__main__":
    main()

def plot_candlestick(stock_code):
    # Fetch historical data
    end_date = datetime.datetime.today()
    start_date = end_date - datetime.timedelta(days=3 * 365)
    stock_data = yf.download(stock_code, start=start_date.strftime('%Y-%m-%d'), end=end_date.strftime('%Y-%m-%d'))

    # Calculate the 200 EMA
    stock_data['EMA_200'] = stock_data['Close'].ewm(span=200, adjust=False).mean()
    stock_data['EMA_50'] = stock_data['Close'].ewm(span=50, adjust=False).mean()
    stock_data['EMA_20'] = stock_data['Close'].ewm(span=20, adjust=False).mean()

    # Filter the last year for display
    stock_data_last_year = stock_data.last('1Y')

    # Filter the data to only the last year for display
    last_year_start_date = end_date - datetime.timedelta(days=200)
    stock_data_last_year = stock_data.loc[last_year_start_date:end_date]

    # Create the candlestick chart using only the last year's data
    fig = go.Figure(data=[go.Candlestick(x=stock_data_last_year.index,
                                         open=stock_data_last_year['Open'],
                                         high=stock_data_last_year['High'],
                                         low=stock_data_last_year['Low'],
                                         close=stock_data_last_year['Close'],
                                         increasing_line_color=' #13abec',
                                         decreasing_line_color='#8a0a0a')])

    # Add the EMA_200 overlay
    fig.update_layout(
    height=600,  # Set the height of the chart
    width=1000,  # Set the width of the chart
    title=f"{stock_code} Stock Price and 200-day EMA",
    yaxis_title='Price (HKD)',
    xaxis_title='Date',
    xaxis_rangeslider_visible=False,  # Hide the range slider
    xaxis_tickformat='%b %Y',  # Set date format to abbreviated month and full year
    plot_bgcolor='white',  # Optional: Set plot background to black
    paper_bgcolor='white',  # Optional: Set paper background to black
    font=dict(color='black')  # Optional: Set font color to white for better contrast
)

# Remove the volume bar chart from the figure if present
# Assuming your figure variable is 'fig'
    if len(fig.data) > 1:
        fig.data = [fig.data[0], fig.data[-1]]  # Keep only the candlestick and EMA line



     # Add the EMAs to the chart
    fig.add_trace(go.Scatter(
        x=stock_data_last_year.index,
        y=stock_data_last_year['EMA_20'],
        mode='lines',
        name='EMA 20',
        line=dict(color='green', width=1.5)
    ))

    fig.add_trace(go.Scatter(
        x=stock_data_last_year.index,
        y=stock_data_last_year['EMA_50'],
        mode='lines',
        name='EMA 50',
        line=dict(color='blue', width=1.5)
    ))

    fig.add_trace(go.Scatter(
        x=stock_data_last_year.index,
        y=stock_data_last_year['EMA_200'],
        mode='lines',
        name='EMA 200',
        line=dict(color='gray', width=1.5)
    ))

    st.plotly_chart(fig)

def main():
    st.sidebar.title("Stock Code")
    index_choice = st.sidebar.selectbox('Select Index', ['HSI', 'HSTECH', 'HSCEI', 'SP 500'], key='index_select')
    
    stock_input = st.sidebar.text_input("Enter a Stock Code:", value="", max_chars=5, key='stock_input')
    
    if stock_input:
        if index_choice == 'SP 500':
            if stock_input.isalpha():
                plot_candlestick(stock_input)  # Ensure this function is defined
            else:
                st.sidebar.error("Please enter a valid stock code for SP 500 stocks.", key='error_sp500')
        else:
            if stock_input.isdigit() and len(stock_input) <= 4:
                formatted_stock_code = stock_input.zfill(4) + ".HK"
                plot_candlestick(formatted_stock_code)  # Ensure this function is defined
            else:
                st.sidebar.error("Please enter a numeric stock code up to 4 digits for Hong Kong stocks.", key='error_hk')

if __name__ == "__main__":
    main()
