from flask import Flask, request, send_file, abort
import pandas as pd
from datetime import timedelta
import io

app = Flask(__name__)


def process_data(df_fuel_transaction, df_delivery_result):
    # Keep only relevant columns and parse dates
    df_fuel_transaction = df_fuel_transaction[['TranDate', 'ทะเบียน']].copy()
    df_fuel_transaction['TranDate'] = pd.to_datetime(
        df_fuel_transaction['TranDate'], format='%d/%m/%Y', errors='coerce'
    )
    df_delivery_result = df_delivery_result[[
        'ออก LDT', 'ลงสินค้า', 'พจส', 'พจส2', 'เลขรถ', 'หัว', 'LDT'
    ]].copy()
    df_delivery_result['ออก LDT'] = pd.to_datetime(
        df_delivery_result['ออก LDT'], format='%d/%m/%Y', errors='coerce'
    )
    df_delivery_result['LDT'] = df_delivery_result['LDT'].astype(str)

    # Initialize new columns
    df_fuel_transaction['พจส'] = None
    df_fuel_transaction['LDT'] = None
    df_fuel_transaction['Medthod'] = None

    # Helper to apply conditions
    def update_df(val_a, val_c, method, condition):
        if pd.isna(val_a) or pd.isna(val_c):
            return
        if condition == 'exact':
            df_filtered = df_delivery_result[
                (df_delivery_result['ออก LDT'] == val_a) &
                (df_delivery_result['หัว'] == val_c)
            ]
        elif condition == 'next_day':
            df_filtered = df_delivery_result[
                (df_delivery_result['ออก LDT'] == val_a + timedelta(days=1)) &
                (df_delivery_result['หัว'] == val_c)
            ]
        elif condition == 'on_or_before':
            df_filtered = df_delivery_result[
                (df_delivery_result['ออก LDT'] < val_a + timedelta(days=1)) &
                (df_delivery_result['หัว'] == val_c)
            ]
        else:
            return

        if not df_filtered.empty:
            unique_names = df_filtered['พจส'].unique()
            LDTs = df_filtered['LDT'].unique()
            joined_ldt = ', '.join(LDTs)
            mask = (
                (df_fuel_transaction['TranDate'] == val_a) &
                (df_fuel_transaction['ทะเบียน'] == val_c)
            )
            if len(unique_names) == 1:
                df_fuel_transaction.loc[mask, 'พจส'] = unique_names[0]
                df_fuel_transaction.loc[mask, 'LDT'] = joined_ldt
                df_fuel_transaction.loc[mask, 'Medthod'] = method
            else:
                joined_names = ', '.join(unique_names)
                df_fuel_transaction.loc[mask, 'พจส'] = joined_names
                df_fuel_transaction.loc[mask, 'LDT'] = joined_ldt
                df_fuel_transaction.loc[mask, 'Medthod'] = 'มีชื่อมากกว่า 1 ในวันเดียว'

    # Apply updates per row
    for val_a, val_c in zip(
        df_fuel_transaction['TranDate'],
        df_fuel_transaction['ทะเบียน']
    ):
        update_df(val_a, val_c, 'TranDate=ออกLDT', 'exact')
        update_df(val_a, val_c, 'เพิ่มวัน', 'next_day')
        update_df(val_a, val_c, 'นับวันย้อนหลัง', 'on_or_before')

    # Truncate LDT at first comma when only one record
    df_fuel_transaction['LDT'] = df_fuel_transaction['LDT'].astype(str)
    mask_multi = df_fuel_transaction['Medthod'] != 'มีชื่อมากกว่า 1 ในวันเดียว'
    df_fuel_transaction.loc[mask_multi, 'LDT'] = (
        df_fuel_transaction.loc[mask_multi, 'LDT']
            .str.partition(',')[0]
            .str.strip()
    )
    return df_fuel_transaction


@app.route('/process', methods=['POST'])
def process_files():
    # Expect two uploaded Excel files
    if 'transaction_file' not in request.files or 'delivery_file' not in request.files:
        return abort(400, 'transaction_file and delivery_file are required')

    trans_file = request.files['transaction_file']
    deliv_file = request.files['delivery_file']

    try:
        df_trans = pd.read_excel(trans_file, sheet_name='รถมีนา')
        df_deliv = pd.read_excel(deliv_file, skiprows=1)
    except Exception as e:
        return abort(400, f'Error reading Excel files: {e}')

    result_df = process_data(df_trans, df_deliv)

    # Write to in-memory buffer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name='result.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    app.run(debug=True)
