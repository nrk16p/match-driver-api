from flask import Flask, request, send_file, abort
import pandas as pd
from datetime import timedelta
import io

app = Flask(__name__)

def process_data(df_fuel_transaction, df_delivery_result):
    # Keep only relevant columns and parse dates
    df_ft = df_fuel_transaction[['TranDate', 'ทะเบียน']].copy()
    df_ft['TranDate'] = pd.to_datetime(
        df_ft['TranDate'], format='%d/%m/%Y', errors='coerce'
    )

    df_dr = df_delivery_result[[
        'ออก LDT', 'ลงสินค้า', 'พจส', 'พจส2', 'เลขรถ', 'หัว', 'LDT'
    ]].copy()
    df_dr['ออก LDT'] = pd.to_datetime(
        df_dr['ออก LDT'], format='%d/%m/%Y', errors='coerce'
    )
    df_dr['LDT'] = df_dr['LDT'].astype(str)

    # Initialize new columns
    df_ft['พจส']    = None
    df_ft['LDT']    = None
    df_ft['Medthod']= None

    def update_df(val_a, val_c, method, condition):
        if pd.isna(val_a) or pd.isna(val_c):
            return
        if condition == 'exact':
            df_f = df_dr[
                (df_dr['ออก LDT'] == val_a) &
                (df_dr['หัว']      == val_c)
            ]
        elif condition == 'next_day':
            df_f = df_dr[
                (df_dr['ออก LDT'] == val_a + timedelta(days=1)) &
                (df_dr['หัว']      == val_c)
            ]
        elif condition == 'on_or_before':
            df_f = df_dr[
                (df_dr['ออก LDT'] < val_a + timedelta(days=1)) &
                (df_dr['หัว']      == val_c)
            ]
        else:
            return

        if not df_f.empty:
            names = df_f['พจส'].unique()
            ldts  = df_f['LDT'].unique()
            joined_ldt = ', '.join(ldts)
            mask = (
                (df_ft['TranDate'] == val_a) &
                (df_ft['ทะเบียน']   == val_c)
            )
            if len(names) == 1:
                df_ft.loc[mask, 'พจส']    = names[0]
                df_ft.loc[mask, 'LDT']    = joined_ldt
                df_ft.loc[mask, 'Medthod']= method
            else:
                df_ft.loc[mask, 'พจส']    = ', '.join(names)
                df_ft.loc[mask, 'LDT']    = joined_ldt
                df_ft.loc[mask, 'Medthod']= 'มีชื่อมากกว่า 1 ในวันเดียว'

    # Apply per-row
    for dt, reg in zip(df_ft['TranDate'], df_ft['ทะเบียน']):
        update_df(dt, reg, 'TranDate=ออกLDT',  'exact')
        update_df(dt, reg, 'เพิ่มวัน',        'next_day')
        update_df(dt, reg, 'นับวันย้อนหลัง', 'on_or_before')

    # Truncate LDT to first comma when single record
    df_ft['LDT'] = df_ft['LDT'].astype(str)
    mask    = df_ft['Medthod'] != 'มีชื่อมากกว่า 1 ในวันเดียว'
    df_ft.loc[mask, 'LDT'] = (
        df_ft.loc[mask, 'LDT']
             .str.partition(',')[0]
             .str.strip()
    )
    return df_ft

@app.route('/process', methods=['POST'])
def process_files():
    # Require both files
    if 'transaction_file' not in request.files or 'delivery_file' not in request.files:
        abort(400, 'transaction_file and delivery_file are required')

    try:
        df_trans = pd.read_excel(request.files['transaction_file'], sheet_name='รถมีนา')
        df_deliv = pd.read_excel(request.files['delivery_file'], skiprows=1)
    except Exception as e:
        abort(400, f'Error reading Excel files: {e}')

    result_df = process_data(df_trans, df_deliv)

    # Return in-memory Excel
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name='result.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
