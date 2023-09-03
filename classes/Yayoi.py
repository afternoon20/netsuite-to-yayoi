import math
import config

class Yayoi:
    # TODO: 余裕があればメンバー変数に設定しておく
    id_flag = None #識別フラグ (1行の伝票データ：2111、複数行の伝票データ1行目：2110、中間の行：2100、最終行：2101)
    slip_number = None #伝票No(管理用)
    settlement_slip = None #決算
    slip_day = None #日付
    debit_name = None #借方勘定科目
    debit_sub_name = None #借方補助科目
    debit_description = None #借方部門
    debit_tax_type = '対象外' #借方税区分
    debit_amount = 0 #借方金額
    debit_tax = 0 #借方税金額
    credit_name = None #貸方勘定科目
    credit_sub_name = None #貸方補助科目
    credit_description = None #貸方部門
    credit_tax_type = '対象外' #貸方税区分
    credit_amount = 0 #貸方金額
    credit_tax = 0 #貸方税金額
    summary = None #摘要
    bill_number = None #番号
    due_date = None #期日
    slip_type = 3 #タイプ（仕訳データの場合は「0」、振伝は「3」）
    system_source = None #生成元
    memo = None #仕訳メモ
    tag1 = None #付箋1
    tag2 = None #付箋2
    adjustment = 'no' #調整（noと記入）

    def set_data(self, transaction):
        self.slip_number = transaction['Document Number']
        self.slip_day = transaction['Date'].strftime('%Y/%m/%d')
        self.summary = transaction['Description']
        self.debit_amount = transaction['Debit'] if transaction['Debit'] else int(0)
        self.credit_amount = transaction['Credit'] if transaction['Credit'] else int (0)
        account = self.trim_for_yayoi(transaction['Account'])
        if (transaction['Debit']):
            self.debit_name = account
        else:
            self.credit_name = account
        self.modify_with_specific_rules()

    @classmethod
    def check_slip_balance(self, yayois, slip_number):
        all_debit_amount = 0
        all_credit_amount = 0
        for yayoi in yayois:
            if yayoi.debit_amount:
                all_debit_amount += math.ceil(yayoi.debit_amount)
            else:
                all_credit_amount += math.ceil(yayoi.credit_amount)

        amount_difference = all_debit_amount - all_credit_amount
        if math.ceil(amount_difference) != 0:
            print('伝票番号：' + slip_number + ' 貸借不一致(' + str(amount_difference) + ')' )

        return

    @classmethod
    def check_slips_balance(self, yayois):
        for slip_number, yayoi_list in yayois.items():
            self.check_slip_balance(yayoi_list, slip_number)

        return

    @classmethod
    def set_id_flag(self, yayois):
        slip_count = len(yayois)
        slip_index = 1
        for yayoi in yayois:
            if slip_count == 1:
                yayoi.id_flag = '2111'
            elif slip_index == 1:
                yayoi.id_flag = '2110'
            elif slip_count == slip_index:
                yayoi.id_flag = '2101'
            else:
                yayoi.id_flag = '2100'
            slip_index += 1

        return

    @classmethod
    def set_id_flags(self,yayois):
        for slip_number, yayoi_list in yayois.items():
            self.set_id_flag(yayoi_list)

        return

    @classmethod
    def set_slip_number_by_index(self, yayois):
        slip_number_index = 1
        for slip_number, yayoi_list in yayois.items():
            for yayoi in yayoi_list:
                yayoi.slip_number = slip_number_index
            slip_number_index += 1

        return

    @classmethod
    def modify_slip_pre_pay_consumption_tax_10_per(self, yayois):
        is_one_slip = False
        contain_tax_slip = False
        is_modified = False
        tax_slip_number = None
        amount_excluding_tax_amount = 0
        if len(yayois) == 1:
            is_one_slip = True
        if not is_one_slip:
            tax_amount = 0
            # TODO:税率定数化、貸方対応
            for index in range(len(yayois)):
                if yayois[index].debit_name in config.local_staticlist.TAX_ACCOUNTS:
                    tax_amount = int(yayois[index].debit_amount)
                    contain_tax_slip = True
                    tax_slip_number = index
                    break
            if contain_tax_slip:
                for yayoi in yayois:
                    if not (yayoi.debit_name in config.local_staticlist.TAX_ACCOUNTS and yayoi.debit_amount == tax_amount):
                        if (int(yayoi.debit_amount or 0) / 10) == tax_amount:
                            yayoi.debit_tax_type = config.staticlist.TAX_NAME_TAXABLE_PURCHASE_10_PER
                            yayoi.debit_amount += tax_amount
                            yayoi.debit_tax = tax_amount
                            yayois.pop(tax_slip_number)
                            is_modified = True
                            break
                if is_modified:
                    # NOTE:再帰的処理
                    self.modify_slip_pre_pay_consumption_tax_10_per(yayois)

                for yayoi in yayois:
                    # NOTE:端数は切り捨て https://www.nta.go.jp/taxes/shiraberu/taxanswer/shohi/6371.htm
                    if not yayoi.debit_name in config.local_staticlist.TAX_ACCOUNTS:
                        amount_excluding_tax_amount += math.floor(int(yayoi.debit_amount or 0) / 10)
                if amount_excluding_tax_amount == tax_amount:
                    for yayoi in yayois:
                        if not yayoi.debit_name in config.local_staticlist.TAX_ACCOUNTS:
                            if int(yayoi.debit_amount or 0) != 0:
                                    yayoi.debit_tax_type = config.staticlist.TAX_NAME_TAXABLE_PURCHASE_10_PER
                            yayoi.debit_tax = math.floor(int(yayoi.debit_amount or 0) * 0.1)
                            yayoi.debit_amount = math.floor(int(yayoi.debit_amount or 0) * 1.1)
                    yayois.pop(tax_slip_number)

        return

    @classmethod
    def modify_slips_pre_pay_consumption_tax_10_per(self, yayois):
        for slip_number, yayoi_list in yayois.items():
            self.modify_slip_pre_pay_consumption_tax_10_per(yayoi_list)

        return

    @classmethod
    def modify_slips_pre_pay_consumption_tax_10_per(self, yayois):
        for slip_number, yayoi_list in yayois.items():
            self.modify_slip_pre_pay_consumption_tax_10_per(yayoi_list)

        return

    @classmethod
    def modify_single_slip(self, yayois):
        for slip_number, yayoi_list in yayois.items():
            self.modify_slip_pre_pay_consumption_tax_10_per(yayoi_list)

        return

    # NOTE:独自ルールは即値で書く
    def modify_with_specific_rules(self):
        # NOTE: 輸出売上
        if (self.credit_name in config.local_staticlist.EXPORT_SALES_ACCOUNT):
            self.credit_tax_type = config.staticlist.TAX_NAME_EXPORT_SALES

    @classmethod
    def trim_for_yayoi(self, account, encoding = 'utf-8'):
        max_byte = 24
        account = str(account)
        while len(account.encode(encoding)) > max_byte:
            account = account[:-1]

        return account.strip()