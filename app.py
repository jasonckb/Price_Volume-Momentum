import streamlit as st
import pandas as pd
import yfinance as yf
import plotly.express as px
import plotly.graph_objs as go
import openpyxl
import datetime

#excel_path = 'C:/Users/user/Desktop/MyScripts/Index Bubble Chart/Index Weight.xlsx'
# Dropbox direct download link
excel_path = 'https://www.dropbox.com/scl/fi/nw5fpges55aff7x5q3gh9/Index-Weight.xlsx?rlkey=rxdopdklplz15jk97zu2sual5&dl=1'

@st.cache(show_spinner=False, allow_output_mutation=True)
def load_data(excel_path, force_reload=False):
    sheet_names = ['HSI', 'HSTECH', 'HSCEI','SP 500']
    dtype = {'Code': str}
    return {name: pd.read_excel(excel_path, sheet_name=name, dtype=dtype) for name in sheet_names}

@st.cache(show_spinner=False, allow_output_mutation=True)
def fetch_and_calculate(df, index_name):
    for index, row in df.iterrows():
        # Adjust the stock code format based on the index
        if index_name == 'SP 500':
            stock_code = row['Code']  # S&P 500 codes should be used as is
        else:
            stock_code = f"{row['Code'].zfill(4)}.HK"  # For Hong Kong stocks, pad and add suffix

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
    st.set_page_config(page_title="Index Constituents Volume & Price Momentum by Jason Chan", layout="wide")
    st.title('Index Components Volume & Price Momentum by Jason Chan')

    if st.sidebar.button('Refresh Data'):
        st.experimental_rerun()

    if 'raw_data' not in st.session_state:
        st.session_state.raw_data = load_data(excel_path)

    processed_data = {name: fetch_and_calculate(st.session_state.raw_data[name].copy(deep=True), name) 
                      for name in st.session_state.raw_data}

    # Store and retrieve the selected index in/from session state
    if 'selected_index' not in st.session_state:
        st.session_state['selected_index'] = 'HSI'

    st.session_state['selected_index'] = st.sidebar.selectbox(
        'Select Index',
        list(processed_data.keys()),
        index=list(processed_data.keys()).index(st.session_state['selected_index'])
    )

    df_display = processed_data[st.session_state['selected_index']].copy(deep=True)

    for name, df in processed_data.items():
        df['Today Pct Change'] = df['Today Pct Change'].apply(format_pct_change)

    # After conversion, calculate the max percentage change
    #max_pct_change = df_display['Today Pct Change'].max()
    #if max_pct_change is not None:
    #    max_pct_change *= 1.1

    
    def color_scale(val):
        try:
        # Convert to float, if conversion fails, it will go to except block
            val = float(val)
        except (ValueError, TypeError):
        # Return a default color if value is not a number or NaN
            return 'gray'
    
    # Check if the value is NaN (after conversion attempt)
        if pd.isna(val):
            return 'gray'

    # Apply color scale based on the value
        if val > 5:
            return 'red'
        elif val > 4:
            return 'Crimson'
        elif val > 3:
            return 'pink'
        elif val > 2:
            return 'brown'
        elif val > 1:
            return 'orange'
        else:
            return 'gray'


    # Clean or preprocess 'Volume Ratio' to handle NaNs and non-numeric values
    df_display['Volume Ratio'] = pd.to_numeric(df_display['Volume Ratio'], errors='coerce').fillna(0)

# Now it's safe to apply 'color_scale'
    df_display['Color'] = df_display['Volume Ratio'].apply(color_scale)


    # Set the y-axis to include negative returns
    min_pct_change = df_display['Today Pct Change'].min()
    max_pct_change = df_display['Today Pct Change'].max() * 1.1  # Adjusted for padding

    fig = px.scatter(df_display, x='Volume Ratio', y='Today Pct Change', size='Weight',
                     hover_data=['Name', 'Code', 'Today Pct Change', 'Volume Ratio'],
                     color='Color', color_discrete_map="identity",
                     title=f'{index_choice} Volume Ratio: Today VS.10 Days Average', height=800, width=1000)

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
        color="white"  # Changed from RebeccaPurple to white for visibility on a black background
    ),
    plot_bgcolor='black',  # Set plot background to black
    paper_bgcolor='black',  # Set overall background to black
    xaxis=dict(
        title_font=dict(size=18, color='white'),  # Ensure title is visible on a black background
        tickfont=dict(size=16, color='white'),  # Ensure ticks are visible
        zeroline=True, zerolinewidth=2, zerolinecolor='white'  # Ensure zero line is visible
    ),
    yaxis=dict(
        title_font=dict(size=18, color='white'),  # Ensure title is visible on a black background
        tickfont=dict(size=16, color='white'),  # Ensure ticks are visible
        zeroline=True, zerolinewidth=2, zerolinecolor='white'  # Ensure zero line is visible
    )
)

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
