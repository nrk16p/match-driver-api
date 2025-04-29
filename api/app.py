from flask import Flask, request, send_file, render_template_string
import pandas as pd
from datetime import timedelta
import io
import time

app = Flask(__name__)

# HTML template embedded for simplicity
HTML = '''
<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Fuel Transaction Processor</title>
  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
  <nav class="navbar navbar-expand-lg navbar-dark bg-primary">
    <div class="container-fluid">
      <a class="navbar-brand" href="#">Fuel Processor</a>
    </div>
  </nav>
  <div class="container py-5">
    <h2 class="mb-4">อัปโหลดไฟล์ Excel เพื่อประมวลผล</h2>
    <form method="post" enctype="multipart/form-data">
      <div class="mb-3">
        <label for="transaction_file" class="form-label">ไฟล์ Fuel Transaction (.xlsx)</label>
        <input class="form-control" type="file" name="transaction_file" id="transaction_file" accept=".xlsx" required>
      </div>
      <div class="mb-3">
        <label for="delivery_file" class="form-label">ไฟล์ Delivery Result (.xlsx)</label>
        <input class="form-control" type="file" name="delivery_file" id="delivery_file" accept=".xlsx" required>
      </div>
      <button type="submit" class="btn btn-primary">Process and Download</button>
    </form>
  </div>
</body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # รับไฟล์จากฟอร์ม
        f1 = request.files.get('transaction_file')
        f2 = request.files.get('delivery_file')
        if not f1 or f1.filename == '' or not f2 or f2.filename == '':
            return render_template_string(HTML + '<div class="container text-danger">กรุณาอัปโหลดไฟล์ทั้งสองให้ครบ!</div>')

        # อ่าน DataFrame
        df_fuel = pd.read_excel(f1, sheet_name="รถมีนา")
        df_fuel = df_fuel[['TranDate', 'ทะเบียน']]
        df_fuel['TranDate'] = pd.to_datetime(df_fuel['TranDate'], format='%d/%m/%Y', errors='coerce')

        df_dr = pd.read_excel(f2, skiprows=1)
        df_dr = df_dr[['ออก LDT', 'ลงสินค้า', 'พจส', 'พจส2', 'เลขรถ', 'หัว', 'LDT']]
        df_dr['ออก LDT'] = pd.to_datetime(df_dr['ออก LDT'], format='%d/%m/%Y', errors='coerce')
        df_dr['LDT'] = df_dr['LDT'].astype(str)

        # สร้างคอลัมน์ใหม่
        df_fuel['พจส'] = None
        df_fuel['LDT'] = None
        df_fuel['Medthod'] = None

        # ฟังก์ชันอัปเดตข้อมูลตามเงื่อนไข
        def update_df(val_a, val_c, method, condition):
            if condition == 'exact':
                filt = (df_dr['ออก LDT'] == val_a) & (df_dr['หัว'] == val_c)
            elif condition == 'next_day':
                filt = (df_dr['ออก LDT'] == val_a + timedelta(days=1)) & (df_dr['หัว'] == val_c)
            else:  # on_or_before
                filt = (df_dr['ออก LDT'] < val_a + timedelta(days=1)) & (df_dr['หัว'] == val_c)

            df_filtered = df_dr[filt]
            if not df_filtered.empty:
                names = df_filtered['พจส'].unique()
                ldts = df_filtered['LDT'].unique()
                joined_ldt = ', '.join(ldts)

                idx = (df_fuel['TranDate'] == val_a) & (df_fuel['ทะเบียน'] == val_c)
                if len(names) == 1:
                    df_fuel.loc[idx, 'พจส'] = names[0]
                    df_fuel.loc[idx, 'LDT'] = joined_ldt
                    df_fuel.loc[idx, 'Medthod'] = method
                else:
                    df_fuel.loc[idx, 'พจส'] = ', '.join(names)
                    df_fuel.loc[idx, 'LDT'] = joined_ldt
                    df_fuel.loc[idx, 'Medthod'] = 'มีชื่อมากกว่า 1 ในวันเดียว'

        # ประมวลผลแต่ละแถว
        for a, c in zip(df_fuel['TranDate'], df_fuel['ทะเบียน']):
            update_df(a, c, 'TranDate=ออกLDT', 'exact')
            update_df(a, c, 'เพิ่มวัน', 'next_day')
            update_df(a, c, 'นับวันย้อนหลัง', 'on_or_before')

        # ตัดค่าหลัง comma ถ้าไม่ใช่กรณีมีชื่อมากกว่า 1
        df_fuel['LDT'] = df_fuel['LDT'].astype(str)
        mask = df_fuel['Medthod'] != 'มีชื่อมากกว่า 1 ในวันเดียว'
        df_fuel.loc[mask, 'LDT'] = df_fuel.loc[mask, 'LDT'].str.partition(',')[0].str.strip()

        # ส่งไฟล์กลับเป็น Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_fuel.to_excel(writer, index=False)
        output.seek(0)
        filename = f"result_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    # GET
    return render_template_string(HTML)

if __name__ == '__main__':
    app.run(debug=True)
