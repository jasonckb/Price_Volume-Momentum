import streamlit as st
import pandas as pd
import yfinance as yf
import plotly.express as px
import plotly.graph_objs as go
import openpyxl
import datetime

# Excel data path
excel_path = 'https://www.dropbox.com/scl/fi/nw5fpges55aff7x5q3gh9/Index-Weight.xlsx?rlkey=rxdopdklplz15jk97zu2sual5&dl=1'

@st.cache_data(show_spinner=False)
def load_data(excel_path):
    sheet_names = ['HSI', 'HSTECH', 'HSCEI', 'SP 500']
    dtype = {'Code': str}
    return {name: pd.read_excel(excel_path, sheet_name=name, dtype=dtype) for name in sheet_names}

def fetch_and_calculate_historical(df, index_name):
    for index, row in df.iterrows():
        stock_code = row['Code'] if index_name == 'SP 500' else f"{row['Code'].zfill(4)}.HK"
        stock = yf.Ticker(stock_code)
        hist = stock.history(period="1mo")  # Fetching 12 days of historical data (including today)

        if len(hist) >= 11:
            yesterday_close = hist['Close'].iloc[-2]
            yesterday_volume = hist['Volume'].iloc[-2]
            avg_volume_10d = hist['Volume'].iloc[-11:-1].mean()  # Calculate average volume for the past 10 days (excluding today)
            yesterday_pct_change = ((yesterday_close - hist['Close'].iloc[-3]) / hist['Close'].iloc[-3]) * 100

            df.at[index, 'Yesterday Close'] = yesterday_close
            df.at[index, 'Today Pct Change'] = round(yesterday_pct_change, 2)
            df.at[index, '10 Day Avg Volume'] = avg_volume_10d
            df.at[index, 'Volume Ratio'] = round(yesterday_volume / avg_volume_10d, 2)
        else:
            # Not enough data, set as None or appropriate default
            df.at[index, 'Yesterday Close'] = None
            df.at[index, 'Today Pct Change'] = None
            df.at[index, '10 Day Avg Volume'] = None
            df.at[index, 'Volume Ratio'] = None

    return df


def fetch_and_calculate_intraday(df, index_name):
    for index, row in df.iterrows():
        try:
            stock_code = row['Code'] if index_name == 'SP 500' else f"{row['Code'].zfill(4)}.HK"
            stock = yf.Ticker(stock_code)
            today_data = stock.history(period="1d")

            if not today_data.empty:
                if 'Yesterday Close' in df.columns and '10 Day Avg Volume' in df.columns:
                    last_close = df.at[index, 'Yesterday Close']
                    avg_volume_10d = df.at[index, '10 Day Avg Volume']

                    today_close = today_data['Close'].iloc[-1]
                    today_volume = today_data['Volume'].iloc[-1]

                    df.at[index, 'Today Pct Change'] = round(((today_close - last_close) / last_close) * 100, 2)
                    df.at[index, 'Volume Ratio'] = round(today_volume / avg_volume_10d, 2)
                else:
                    print(f"Missing required columns for {stock_code}")
            else:
                print(f"No data available for {stock_code}")
        except Exception as e:
            print(f"Error processing {stock_code}: {e}")
    
    return df

def color_scale(val):
    if val > 5: return 'red'
    elif val > 4: return 'Crimson'
    elif val > 3: return 'pink'
    elif val > 2: return 'brown'
    elif val > 1: return 'orange'
    else: return 'gray'

def generate_plot(df_display, index_choice):
    if 'Today Pct Change' in df_display.columns:
        min_pct_change = df_display['Today Pct Change'].min() * 1.1 if pd.notna(df_display['Today Pct Change']).any() else -10
        max_pct_change = df_display['Today Pct Change'].max() * 1.1 if pd.notna(df_display['Today Pct Change']).any() else 10
    else:
        min_pct_change = -10
        max_pct_change = 10
    
    fig = px.scatter(
        df_display, 
        x='Volume Ratio', 
        y='Today Pct Change', 
        size='Weight',
        hover_data=['Name', 'Code'],
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
        yaxis=dict(title_font=dict(color='white'), tickfont=dict(color='white'), range=[min_pct_change, max_pct_change])
    )

    if df_display['Volume Ratio'].min() > 0:
        fig.update_xaxes(type='log')
    else:
        fig.update_xaxes(type='linear')

    return fig



def main():
    st.set_page_config(page_title="Index Constituents Volume & Price Momentum", layout="wide")
    st.title('Index Components Volume & Price Momentum')

    if 'raw_data' not in st.session_state:
        st.session_state['raw_data'] = load_data(excel_path)
    
    index_choice = st.sidebar.selectbox('Select Index', ['HSI', 'HSTECH', 'HSCEI', 'SP 500'])

    if 'processed_data' not in st.session_state or st.sidebar.button('Daily Historical Update'):
        st.session_state['processed_data'] = {
            name: fetch_and_calculate_historical(st.session_state['raw_data'][name].copy(), name) 
            for name in st.session_state['raw_data']
        }

    if st.sidebar.button('Intraday Refresh'):
        st.session_state['processed_data'][index_choice] = fetch_and_calculate_intraday(
            st.session_state['processed_data'][index_choice].copy(), index_choice)

    df_display = st.session_state['processed_data'].get(index_choice, pd.DataFrame()).copy()
    if not df_display.empty:
        df_display['Color'] = df_display['Volume Ratio'].apply(color_scale)
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
