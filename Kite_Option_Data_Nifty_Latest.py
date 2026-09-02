import pandas as pd
import datetime as dt
from datetime import datetime, timedelta
from kiteapp import KiteApp
from kiteconnect import KiteTicker
from time import sleep
from py_vollib.black_scholes import black_scholes
from py_vollib.black_scholes.implied_volatility import implied_volatility
from py_vollib.black_scholes.greeks.analytical import delta, theta, gamma, vega, rho
import xlwings as xw
import threading
from threading import Thread
from queue import Queue, Empty
import pythoncom
import logging.config
import numpy as np
import asyncio

# Configure logging
logging.basicConfig(filename='script_log.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

logging.info("Script started")

# Function to check if Excel application is ready
def is_excel_ready():
    try:
        app = xw.App(visible=False)
        app.kill()
        return True
    except Exception as e:
        logging.error(f"Excel is not ready: {e}")
        return False

# Function to open the Excel workbook with retries
def open_excel_workbook(file_name, retries=3, delay=5):
    attempt = 0
    while attempt < retries:
        if is_excel_ready():
            try:
                wb = xw.Book(file_name)
                logging.info("Workbook opened successfully")
                return wb
            except FileNotFoundError:
                logging.error(f"No such file: '{file_name}'")
                return None
            except Exception as e:
                logging.error(f"Attempt {attempt + 1} - An error occurred while opening the Excel file: {e}")
        else:
            logging.error(f"Attempt {attempt + 1} - Excel is not ready. Retrying in {delay} seconds...")
        
        attempt += 1
        sleep(delay)
    return None

# Open the Excel workbook
file_name = "tick_data1_Nifty-Shahid.xlsm"
wb = open_excel_workbook(file_name)

# Check if the workbook was opened successfully
if not wb:
    logging.critical(f"Failed to open the file '{file_name}' after multiple attempts")
    raise FileNotFoundError(f"Failed to open the file '{file_name}' after multiple attempts")

# Print out the sheet names to confirm access
logging.info(f"Sheet names in the workbook: {[sheet.name for sheet in wb.sheets]}")

# Access sheets
try:
    ws_all_inst = wb.sheets['All_Instruments']
    ws_TickData = wb.sheets['TickData']
    ws_Expiry = wb.sheets['Underline']
    ws_Config = wb.sheets['Config']
    logging.info("Sheets accessed successfully")
except Exception as e:
    logging.critical(f"Error accessing sheets: {e}")
    raise e

# Load Kite API credentials from Excel
try:
    enctoken_cell = ws_Config.range('B1').value
    if enctoken_cell is None:
        logging.critical("Cell B1 in Config sheet is empty. Please provide the enctoken value.")
        raise ValueError("Cell B1 in Config sheet is empty. Please provide the enctoken value.")
    enctoken = enctoken_cell.strip()
    logging.info("Kite API credentials loaded successfully")
except Exception as e:
    logging.critical(f"Error loading Kite API credentials: {e}")
    raise e

# Initialize Kite App and WebSocket
kite = KiteApp("s1", "SJ8828", enctoken)
kws = kite.kws()


# Read the instrument name and segment from the Config sheet
instrument_name = ws_Config['B2'].value
instrument_segment = ws_Config['B3'].value

instrument_name_fut = ws_Config['C2'].value
instrument_segment_fut = ws_Config['C3'].value

instrument_name_indice = ws_Config['D2'].value
instrument_segment_indice = ws_Config['D3'].value

# Fetch all instruments
all_instruments = kite.instruments()
df = pd.DataFrame(all_instruments)

nfo_instruments = df[df['segment'] == instrument_segment]
nfo_instruments = nfo_instruments[nfo_instruments['name'] == instrument_name]

nfo_fut = df[df['segment'] == instrument_segment_fut]
nfo_fut = nfo_fut[nfo_fut['name'] == instrument_name_fut]

inst_indice = df[df['segment'] == instrument_segment_indice]
inst_indice = inst_indice[inst_indice['name'] == instrument_name_indice]

# Assuming nfo_fut and inst_indice are already defined dataframes
merged_unerline_df = pd.concat([nfo_fut, inst_indice])

#ws_all_inst['A1'].value = df

# Extract and sort the unique expiry dates
exp = sorted(set(nfo_instruments['expiry'].to_list()))
# Convert the list of expiry dates to a DataFrame
exp_df = pd.DataFrame(exp, columns=['All Expiries'])
# Display the DataFrame without the index
Print_Exp = exp_df.to_string(index=False)

# Write the header
ws_Config['S1'].value = 'All Expiries'
# Write each expiry date to a new row
for idx, date in enumerate(exp, start=2):
    ws_Config[f'S{idx}'].value = date

def get_weekly_expiry(df, index):
    df_filtered = df[(df['name'] == index) & (df['segment'] == instrument_segment)]
    exp = sorted(set(df_filtered['expiry'].to_list()))
    td = dt.date.today()
    nearest_expiry = min(exp, key=lambda x: (x - td))
    exp.remove(nearest_expiry)
    second_nearest_expiry = min(exp, key=lambda x: (x - td)) if exp else None
    if second_nearest_expiry:
        df_result = df_filtered[df_filtered['expiry'].isin([nearest_expiry, second_nearest_expiry])]
    else:
        df_result = df_filtered[df_filtered['expiry'] == nearest_expiry]
    return df_result, nearest_expiry, second_nearest_expiry

df, nearest_expiry, second_nearest_expiry = get_weekly_expiry(df, instrument_name)
print('Nearest Expiry:', nearest_expiry)
print('Second Nearest Expiry:', second_nearest_expiry)

underline_tokens = merged_unerline_df['instrument_token'].to_list()
# Define the new token to be added
index_token = 256265

# Append the new token to the list
underline_tokens.append(index_token)

# Print the updated list
print(underline_tokens)


value_in_ai4 = ws_Expiry['B9'].value

# Read the number of strikes from cell B4 in the Config sheet
number_of_strikes = ws_Config.range('B4').value
# Read the number of strikes from cell B4 in the Config sheet
number_of_oi_strikes = ws_Config.range('C4').value
# Calculate the step size (assuming each strike is a multiple of 50)
step_size = 50
# Calculate the range based on the number of strikes
range_value = (number_of_strikes // 2) * step_size
# Calculate the range based on the number of strikes fro OI
range_value_oi = (number_of_oi_strikes // 2) * step_size

# Calculate the lower and upper strike prices
lower_strike = int(value_in_ai4 - range_value)
upper_strike = int(value_in_ai4 + range_value)

# Calculate the lower and upper strike prices
lower_oi_strike = int(value_in_ai4 - range_value_oi)
upper_oi_strike = int(value_in_ai4 + range_value_oi)

# Print the results
print(f"Lower Strike: {lower_strike}")
print(f"Upper Strike: {upper_strike}")
# Print the results
print(f"Lower OI Strike: {lower_oi_strike}")
print(f"Upper OI Strike: {upper_oi_strike}")

# Filter the DataFrame based on the strike price range
filtered_df = df[(df['strike'] >= lower_strike) & (df['strike'] <= upper_strike)]
ltp_tokens = filtered_df['instrument_token'].tolist()

# Filter the DataFrame based on the strike price range
filtered_df = df[(df['strike'] >= lower_oi_strike) & (df['strike'] <= upper_oi_strike)]
oi_tokens = filtered_df['instrument_token'].tolist()

ws_Config['L1'].value = 'All Tokens'
# Write each expiry date to a new row
for idx, date in enumerate(ltp_tokens, start=2):
    ws_Config[f'L{idx}'].value = date

index_last_price = None
index_open_price = None
index_close_price = None



def merge_dataframes(df_instruments, df_tick_data, instrument_columns, tick_data_columns):
    df_instruments = df_instruments[instrument_columns]
    df_tick_data = df_tick_data[tick_data_columns]
    merged_df = pd.merge(df_instruments, df_tick_data, on='instrument_token', how='inner')
    return merged_df

def remain_days(expiry):
    return (pd.to_datetime(pd.to_datetime(expiry).strftime("%Y-%m-%d 15:30:00")) - dt.datetime.now()) / dt.timedelta(days=1) / 365

def remain_days_open(expiry):
    # Convert the expiry to a datetime object
    expiry_date = pd.to_datetime(expiry).strftime("%Y-%m-%d 15:30:00")
    expiry_date = pd.to_datetime(expiry_date)
    
    # Set the current time to 9:16 AM of the current day
    current_time = dt.datetime.now().replace(hour=9, minute=15, second=5, microsecond=0)
    
    # Calculate the difference in days
    remaining_days = (expiry_date - current_time) / dt.timedelta(days=1) / 365
    return remaining_days

def calculate_iv_and_greeks(row):
    global index_last_price
    global index_open_price
    
    # Calculate values based on last price
    last_price = row['last_price']
    if last_price == 0:
        last_price_values = [float('nan')] * 5
    else:
        F_last = index_last_price if index_last_price is not None else 22980
        K = int(row['strike'])
        r = row['remain_days']
        t = 0.1
        flag = 'c' if row['instrument_type'] == 'CE' else 'p'
        try:
            iv_last = implied_volatility(last_price, F_last, K, r, t, flag)
            dlt_last = delta(flag, F_last, K, r, t, iv_last)
            th_last = round(theta(flag, F_last, K, r, t, iv_last), 3)
            gm_last = round(gamma(flag, F_last, K, r, t, iv_last), 5)
            vg_last = round(vega(flag, F_last, K, r, t, iv_last), 3)
            last_price_values = [round((iv_last * 100), 2), round((dlt_last), 2), th_last, gm_last, vg_last]
        except Exception as e:
            last_price_values = [float('nan')] * 5

    # Calculate values based on open price
    open_price = row['open']
    if open_price == 0 or 'remain_days_open' not in row:
        open_price_values = [float('nan')] * 5
    else:
        F_open = index_open_price if index_open_price is not None else 22980
        r_open = row['remain_days_open']
        try:
            iv_open = implied_volatility(open_price, F_open, K, r_open, t, flag)
            dlt_open = delta(flag, F_open, K, r_open, t, iv_open)
            th_open = round(theta(flag, F_open, K, r_open, t, iv_open), 3)
            gm_open = round(gamma(flag, F_open, K, r_open, t, iv_open), 5)
            vg_open = round(vega(flag, F_open, K, r_open, t, iv_open), 3)
            open_price_values = [round((iv_open * 100), 2), round((dlt_open), 2), th_open, gm_open, vg_open]
        except Exception as e:
            open_price_values = [float('nan')] * 5

    return pd.Series(last_price_values + open_price_values)

# Simplified flatten data without depth buy/sell fields
def flatten_data(tick, source='TickData'):
    return {
        'instrument_token': tick.get('instrument_token'),
        'last_price': tick.get('last_price', 0),
        'average_traded_price': tick.get('average_traded_price', 0),
        'volume_traded': tick.get('volume_traded', 0),
        'total_buy_quantity': tick.get('total_buy_quantity', 0),
        'total_sell_quantity': tick.get('total_sell_quantity', 0),
        'open': tick.get('ohlc', {}).get('open', 0),
        'high': tick.get('ohlc', {}).get('high', 0),
        'low': tick.get('ohlc', {}).get('low', 0),
        'close': tick.get('ohlc', {}).get('close', 0),
        'change': round(tick.get('change', 0), 2),
        'oi': tick.get('oi', 0),
        'oi_day_high': tick.get('oi_day_high', 0),
        'oi_day_low': tick.get('oi_day_low', 0),
        'exchange_timestamp': tick.get('exchange_timestamp', 0),
        'source': source
    }

# Underline data remains unchanged
def flatten_underline_data(tick, source='Underline'):
    return {
        'instrument_token': tick.get('instrument_token'),
        'last_price': tick.get('last_price', 0),
        'volume_traded': tick.get('volume_traded', 0),
        'open': tick.get('ohlc', {}).get('open', 0),
        'high': tick.get('ohlc', {}).get('high', 0),
        'low': tick.get('ohlc', {}).get('low', 0),
        'close': tick.get('ohlc', {}).get('close', 0),
        'change': round(tick.get('change', 0), 2),
        'oi': tick.get('oi', 0),
        'exchange_timestamp': tick.get('exchange_timestamp', 0),
        'source': source
    }
# WebSocket handlers
def on_ticks(ws, ticks):
    global index_last_price, index_open_price, index_close_price
    for tick in ticks:
        if tick['instrument_token'] == index_token:
            index_last_price = tick['last_price']
            index_close_price = tick.get('ohlc', {}).get('close', 0)
            index_open_price = tick.get('ohlc', {}).get('open', 0)

        flattened_tick = flatten_data(tick)
        asyncio.run_coroutine_threadsafe(data_queue.put(flattened_tick), event_loop)
        if tick['instrument_token'] in underline_tokens:
            flattened_underline_tick = flatten_underline_data(tick)
            asyncio.run_coroutine_threadsafe(underline_data_queue.put(flattened_underline_tick), event_loop)

def on_connect(ws, response):
    ws.subscribe(ltp_tokens)
    ws.set_mode(ws.MODE_FULL, ltp_tokens)
    ws.subscribe(underline_tokens)
    ws.set_mode(ws.MODE_FULL, underline_tokens)
    print('Connected and Subscribed')

def on_close(ws, code, reason):
    print(f"Connection closed: {code} - {reason}")
    reconnect(ws)

def on_error(ws, code, reason):
    print(f"Connection error: {code} - {reason}")
    reconnect(ws)

def reconnect(ws):
    sleep(5)
    ws.connect()

# Function to fetch historical data for a given date and token
def fetch_historical_data(token, date_str):
    try:
        historical_data = kite.historical_data(token, date_str, date_str, 'day', oi=1)
        if not historical_data:
            print(f"No data for token {token} on: {date_str}")
            return None
        return historical_data
    except Exception as e:
        print(f"Error fetching data for token {token} on {date_str}: {e}")
        return None
async def data_processor():
    token_row_map = {}
    df_tick_data = pd.DataFrame()

    # Load previous trading date
    previous_day_str = ws_Config.range('B6').value
    previous_day = pd.to_datetime(previous_day_str)
    previous_day_str = previous_day.strftime('%Y-%m-%d')
    
    # Fetch historical data
    historical_data_combined = []
    for token in ltp_tokens:
        hist = fetch_historical_data(token, previous_day_str)
        if hist:
            for rec in hist:
                rec['instrument_token'] = token
            historical_data_combined.extend(hist)
        sleep(0.05)
    historical_df = pd.DataFrame(historical_data_combined)
    historical_df['date'] = pd.to_datetime(historical_df['date'])

    while True:
        batch_ticks = []
        while not data_queue.empty():
            batch_ticks.append(await data_queue.get())

        if batch_ticks:
            batch_df = pd.DataFrame(batch_ticks)
            for tick in batch_ticks:
                tok = tick['instrument_token']
                if tok not in token_row_map:
                    token_row_map[tok] = len(df_tick_data)
                    df_tick_data = pd.concat([df_tick_data, pd.DataFrame([tick])], ignore_index=True)
                else:
                    df_tick_data.loc[token_row_map[tok]] = tick

            instrument_cols = ['instrument_token', 'tradingsymbol', 'expiry', 'strike', 'instrument_type']
            tick_cols = [
                'instrument_token', 'last_price', 'average_traded_price', 'volume_traded',
                'total_buy_quantity', 'total_sell_quantity', 'open', 'high', 'low', 'close',
                'change', 'oi', 'oi_day_high', 'oi_day_low', 'exchange_timestamp'
            ]
            merged = merge_dataframes(nfo_instruments, df_tick_data, instrument_cols, tick_cols)
            merged['remain_days'] = merged['expiry'].apply(remain_days)
            merged['remain_days_open'] = merged['expiry'].apply(remain_days_open)
            merged[['iv', 'delta', 'theta', 'gamma', 'vega', 'open iv', 'open delta', 'open theta', 'open gamma', 'open vega']] = merged.apply(calculate_iv_and_greeks, axis=1)
            merged = pd.merge(merged, historical_df, on='instrument_token', suffixes=('', '_prev'))

            await excel_update_queue.put(merged.copy())
            print("Data Processor: Processed data to DataFrame")

        await asyncio.sleep(0.4)


async def underline_data_processor():
    df_underline_data = pd.DataFrame(columns=['instrument_token', 'last_price', 'volume_traded', 'open', 'high', 'low', 'close', 'change', 'oi', 'exchange_timestamp'])
    token_row_map = {}

    while True:
        batch_ticks = []
        while not underline_data_queue.empty():
            tick = await underline_data_queue.get()
            batch_ticks.append(tick)

        if batch_ticks:
            batch_df = pd.DataFrame(batch_ticks)
            batch_df = batch_df.dropna(how='all')
            if not batch_df.empty:
                for tick in batch_ticks:
                    token = tick['instrument_token']
                    if token not in token_row_map:
                        row_index = len(df_underline_data)
                        df_underline_data = pd.concat([df_underline_data, pd.DataFrame([tick])], ignore_index=True)
                        token_row_map[token] = row_index
                    else:
                        row_index = token_row_map[token]
                        df_underline_data.loc[row_index] = tick

                # Filter nfo_fut by underline_tokens
                nfo_fut_filtered = merged_unerline_df[merged_unerline_df['instrument_token'].isin(underline_tokens)]
                
                # Merge with nfo_fut_filtered
                instrument_columns = ['instrument_token', 'tradingsymbol', 'name', 'expiry']
                merged_underline_data = pd.merge(df_underline_data, nfo_fut_filtered[instrument_columns], on='instrument_token', how='inner')

                await underline_excel_update_queue.put(merged_underline_data.copy())
                print("Underline Data Processor: Processed data to DataFrame")

        await asyncio.sleep(0.4)

async def excel_update_worker():
    pythoncom.CoInitialize()
    try:
        workbook = xw.Book("tick_data1_Nifty-Shahid.xlsm")
        ws_TickData = workbook.sheets["TickData"]

        while True:
            if not excel_update_queue.empty():
                df_TickData = await excel_update_queue.get()
                try:
                    ws_TickData.range('A1').options(index=False).value = df_TickData
                    print("Update Excel: Updated TickData sheet")
                except Exception as e:
                    print("An error occurred while updating Excel:")
                    print(f"Error: {e}")
            await asyncio.sleep(0.4)
    finally:
        pythoncom.CoUninitialize()

async def underline_excel_update_worker():
    pythoncom.CoInitialize()
    try:
        workbook = xw.Book("tick_data1_Nifty-Shahid.xlsm")
        ws_Underline = workbook.sheets["Underline"]

        while True:
            if not underline_excel_update_queue.empty():
                df_underline = await underline_excel_update_queue.get()
                try:
                    ws_Underline.range('A1').options(index=False).value = df_underline
                    print("Update Excel: Updated Underline sheet")
                except Exception as e:
                    print("An error occurred while updating Excel:")
                    print(f"Error: {e}")
                try:
                    ws_Underline.range('A10').value = "Last Price"
                    ws_Underline.range('B10').value = "Open Price"
                    ws_Underline.range('C10').value = "Close Price"
                    ws_Underline.range('A11').value = index_last_price
                    ws_Underline.range('B11').value = index_open_price
                    ws_Underline.range('C11').value = index_close_price
                except Exception as e:
                    print("An error occurred while updating index prices:")
                    print(f"Error: {e}")

            await asyncio.sleep(0.4)
    finally:
        pythoncom.CoUninitialize()

def logging_and_profiling():
    pass

# Initialize queues
data_queue = asyncio.Queue()
underline_data_queue = asyncio.Queue()
excel_update_queue = asyncio.Queue()
underline_excel_update_queue = asyncio.Queue()

# WebSocket event handlers
kws.on_ticks = on_ticks
kws.on_connect = on_connect
kws.on_close = on_close
kws.on_error = on_error

# Function to run WebSocket
def run_ws(loop):
    asyncio.set_event_loop(loop)
    kws.connect(threaded=True)
    loop.run_forever()

# Create an event loop and start WebSocket
event_loop = asyncio.new_event_loop()
ws_thread = threading.Thread(target=run_ws, args=(event_loop,))
ws_thread.start()

# Run async tasks on the main thread's event loop
asyncio.ensure_future(data_processor())
asyncio.ensure_future(underline_data_processor())
asyncio.ensure_future(excel_update_worker())
asyncio.ensure_future(underline_excel_update_worker())

# Start logging and profiling in a separate thread
logging_and_profiling_thread = threading.Thread(target=logging_and_profiling)
logging_and_profiling_thread.start()

print("Start...")

try:
    asyncio.get_event_loop().run_forever()
except KeyboardInterrupt:
    kws.close()
    ws_thread.join()
    event_loop.stop()
    logging_and_profiling_thread.join()
    logging.info("Script terminated by user")
