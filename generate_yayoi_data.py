import util
from classes.Yayoi import Yayoi as Yayoi

# python generate_yayoi_data.py

wb = util.netsuite.load_general_ledger_work_book()
ws = wb.worksheets[0]
# TODO: 項目名を定数化
header = {'Account': None, 'Date': None, 'Document Number': None, 'Description': None, 'Debit': None, 'Credit': None}
start_row = util.netsuite.get_start_row(ws, 'Account')
util.netsuite.set_header_row(ws, header, start_row)

if (None in header.values()):
    print('必要な項目が不足しています。')
    exit
else:
    util.netsuite.set_account(ws, header['Account'], start_row + 1)
    transactions = []
    for row in ws.iter_rows(start_row + 1):
        transaction = {}
        for cell in row:
            if (cell.column in header.values() and ws.cell(cell.row, header['Date']).value not in ['', 0, None, False]):
                key = [key for key, value in header.items() if value == cell.column][0]
                transaction[key] = cell.value
        if (transaction):
            # NOTE:伝票番号が空欄の行は伝票番号に日付+blankを付与
            if transaction['Document Number'] in ['', 0, None, False]:
                transaction['Document Number'] = transaction['Date'].strftime('%Y%m%d') + 'blank'
            transactions.append(transaction)
    transactions = sorted(transactions, key=lambda transaction: (
        transaction['Date'], transaction['Document Number']))

    yayois = util.yayoi.generate_yayois(transactions)
    # TODO: 元のDOCUMENT NUMBERを把握するため、実行時の引数で制御したい
    Yayoi.set_slip_number_by_index(yayois)
    Yayoi.modify_slips_pre_pay_consumption_tax_10_per(yayois)
    Yayoi.check_slips_balance(yayois)
    Yayoi.set_id_flags(yayois)
    util.yayoi.create_yayoi_import_file(yayois)
