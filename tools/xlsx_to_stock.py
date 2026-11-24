"""ฝังข้อมูลจากไฟล์ LAY OUT 02 (.xlsx) ลงใน stock.html

ใช้:  python tools/xlsx_to_stock.py "ไฟล์.xlsx" [stock.html]

- อ่านชีต "ใบจดสต๊อค" (แถว 4 ลงไป: รหัส, สินค้า, จำนวน, หน่วย, โลเคชั่น, หมายเหตุ)
  และชีต "สินค้า" (MATERIAL CODE, DESCRIPTION, UNIT)
- อ่านค่าที่ Excel คำนวณไว้แล้ว (data_only) ไฟล์ต้องถูกเซฟจาก Excel มาก่อน
- แทนที่เฉพาะช่วงระหว่าง // @@DATA-START กับ // @@DATA-END ใน stock.html
"""
import datetime
import json
import sys
import warnings

import openpyxl

warnings.filterwarnings('ignore')


def clean(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, (int, float)):
        return round(v, 4)
    return str(v).strip()


def main(xlsx, html='stock.html'):
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)

    stock = []
    for r in wb['ใบจดสต๊อค'].iter_rows(min_row=4, max_col=6, values_only=True):
        if r[0] is None or str(r[0]).strip() == '':
            continue
        stock.append([clean(x) for x in r])

    products = []
    for r in wb['สินค้า'].iter_rows(min_row=2, max_col=3, values_only=True):
        if r[0] is None or str(r[0]).strip() == '':
            continue
        products.append([clean(x) for x in r])

    if not stock:
        sys.exit('ไม่พบข้อมูลในชีต ใบจดสต๊อค — หยุด ไม่แก้ไฟล์')
    if any(row[2] == '' for row in stock):
        print('คำเตือน: มีแถวที่จำนวนว่าง (ไฟล์อาจยังไม่ถูกคำนวณใน Excel)')

    d = datetime.date.today()
    date_th = f'{d.day}/{d.month}/{d.year + 543}'
    js = lambda o: json.dumps(o, ensure_ascii=False, separators=(',', ':'))
    block = ('// @@DATA-START\n'
             f"const DATA_DATE = {js(date_th)};\n"
             f'const STOCK = {js(stock)};     // [รหัส, สินค้า, จำนวน, หน่วย, โลเคชั่น, หมายเหตุ] จากชีต "ใบจดสต๊อค"\n'
             f'const PRODUCTS = {js(products)};  // [รหัส, สินค้า, หน่วย] จากชีต "สินค้า"\n'
             '// @@DATA-END')

    src = open(html, encoding='utf-8').read()
    a, b = src.index('// @@DATA-START'), src.index('// @@DATA-END') + len('// @@DATA-END')
    open(html, 'w', encoding='utf-8').write(src[:a] + block + src[b:])
    print(f'ok: ใบจดสต๊อค {len(stock)} แถว, สินค้า {len(products)} รหัส → {html}')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(*sys.argv[1:3])
