# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
#%% IMPORT LIBRARIES
import path,pandas,numpy,matplotlib,os,re,pathlib,jdatetime, xlrd , asposecells ,xlwings, os
# jpype.startJVM()
# from asposecells.api import Workbook

import variables


#%% SETTINGS
pandas.options.display.float_format = '{:.2f}'.format
    
directory = os.getcwd()


#%% FUNCTIONS

def currency_formart(x,reverse=False):
    if not reverse:
        return "{:,}".format(int(x))
    return int(x.replace(',',''))
    
def is_nan(x):
    return x!=x


def PERSIAN_DIC_DF(dic):
    per_dic = {}
    for key in list(dic.keys()):
        per_dic[variables.GET_PERSIAN_NAME(key)] = PERSIAN_OF_DF(dic[key])
    
    return per_dic
    

        
def PERSIAN_OF_DF(eng_df):
    per_df_columns =[]
    for col in eng_df.columns:
        per_df_columns.append(variables.GET_PERSIAN_NAME(col))
    
    per_df = pandas.DataFrame(columns=per_df_columns)
    for col in eng_df:
        colobj = eng_df[col]
        
    for i in range(len(eng_df)):
        item = eng_df.loc[i]
        temp = []
        
        for col in list(eng_df.columns):
            if type (item[col]) is dict:
                temp.append(item[col])

            else:
                temp.append(variables.GET_PERSIAN_NAME(item[col]))
            
        per_df.loc[len(per_df)] = temp
            
    return per_df
            
        
    

def autofit(file_path,pdf_save=False):  
    with xlwings.App(visible=False):
        book = xlwings.Book(file_path)
        for sheet in book.sheets:
        # sheet = book.sheets[0]
            sheet.autofit()
        book.save(file_path)
        
        if pdf_save:
            sheet = book.sheets[-1]
            sheet.api.PageSetup.Orientation = xlwings.constants.PageOrientation.xlLandscape
            book.to_pdf(file_path[:file_path.index('.')]+'.pdf', include=book.sheet_names[-1])
        
        book.close()

def CHANGE_COLUMN_TO_CURRENCY_FORMAT(df,column_names):
    for i in range(len(df)):
        for col_name in column_names:
            df.at[i,col_name] = currency_formart(df.loc[i][col_name])
    return df


def SAVE_DFDICT_TO_EXCELL_SHEETS(dfs_dict, xlsx_path,pdf_save=False):
    # with pandas.ExcelWriter(xlsx_path,engine="openpyxl") as writer:
    writer = pandas.ExcelWriter(xlsx_path) 
    a = list(dfs_dict.keys())
    a.sort()
    for sheet_name in a:
        df = dfs_dict[sheet_name]
        
        df = df.style.set_properties(**{
        'font-size': '11pt',
        'text-align': 'center'})
        
        df.to_excel(writer, sheet_name=sheet_name, index=False, na_rep='')
        data = df.data
        if not data.empty:
            for column in data:
                column_length = max(data[column].astype(str).map(len).max(), len(column))
                col_idx = data.columns.get_loc(column)
                writer.sheets[sheet_name].set_column(col_idx, col_idx, column_length)
        
    writer.close()    
            
    if pdf_save:
        df = dfs_dict[list(dfs_dict.keys())[-1]]
        
        df = df.style.set_properties(**{
        'font-size': '16pt',
        'border-bottom': '1pt solid gray'
        })
        # df = df.style.background_gradient(axis=None,vmin=1, vmax=5, cmap="YlGnBu")
        
        html_path = str(xlsx_path)[:str(xlsx_path).index('.')]+'.html'
        pdf_path = str(xlsx_path)[:str(xlsx_path).index('.')]+'.pdf'
        
        f = open(html_path ,'w')
        a = df.to_html()
        with open(html_path, "w", encoding="utf-8") as file:
            file.writelines('<meta charset="UTF-8">\n')
            file.write(a)
        

def ASSIGN_SIZE_COLUMN(trade_df):
    size_column= []
    
    for i in range(len(trade_df)):
        length = str(trade_df.loc[i]['length'])
        width = str(trade_df.loc[i]['width'])
        if length[-2:] == '.0':
            length = str(int(trade_df.loc[i]['length']))
        if width[-2:] == '.0':
            width = str(int(trade_df.loc[i]['width']))
        
        if width.isdigit():
            size_column.append(length + ' * ' + width)
            
        else:
            size_column.append('')
    
    trade_df = trade_df.assign(size=size_column)
    
    return trade_df


def GET_MONTHES_FOLDERS():
    monthes_list =[]
    temp = next(os.walk('.'))[1]
    for item in temp:
        if item[0].isdigit() and 'total' not in item:
            monthes_list.append(pathlib.Path(directory+'/' + item))
    monthes_list.sort()
    
    return monthes_list


def GET_MONTH_TRADES_DF_DIC(month_folder,buy_sell):
    for file in os.listdir(month_folder):
        if file[:2].isdigit() and buy_sell in file:
            file_path = pathlib.Path(str(month_folder)+'/' + file)
    file = pandas.ExcelFile(file_path)
    sheet_names = file.sheet_names
    month_df_dict = {}
    for sheet_name in sheet_names:
        month_df_dict[sheet_name] = file.parse()
        if sheet_name == 'sheet1':
            month_df_dict[sheet_name].dropna(subset=['date'] ,inplace=True)
                    
            month_df_dict[sheet_name]['fee'].fillna(0 ,inplace=True)
            
            
    
    return month_df_dict

def ASSIGN_FULL_NAME_COLUMN(df,col1_name,col2_name):
    new_column= []
    for i in range(len(df)):
        col1 = df.loc[i][col1_name]
        col2 = df.loc[i][col2_name]
        
        if is_nan(col1):
            col1 = ''
        if is_nan(col2):
            col2 = ''
            
        new_column.append(str(col1) + ' ' + str(col2))
    
    _df = df.assign(full_name=new_column)
    
    return _df


def SUMMERIZE_MONTH_STONE_TRADES_DF(month_trade_df,trade_type):
    side = 1
    if 'buy' in trade_type:    
        month_report_df_columns = ['stone_type' ,'stone_name' ,'sizes' ,'meterage_total' ,'meterage_details' ,
                                    'paid_total' ,'paid_details' , 'average_buy_price']
        trader_type = 'seller'
        
        
    elif 'sell' in trade_type:
        month_report_df_columns = ['stone_type' ,'stone_name' ,'sizes' ,'meterage_total' ,'meterage_details' ,
                                    'paid_total' ,'paid_details' , 'average_sell_price']
        trader_type = 'buyer'
        
    elif 'both' in trade_type:
        month_report_df_columns = ['stone_type' ,'stone_name' ,'sizes' ,
                                   'buy_meterage_total', 'buy_meterage_details', 'paid_total', 'paid_details', 'average_buy_price',
                                   'sell_meterage_total', 'sell_meterage_details', 'earned_total', 'earned_details', 'average_sell_price' ]
        
        trader_type = ['seller','buyer']
    
    month_report_df = pandas.DataFrame(columns=month_report_df_columns)
    
    month_trade_df = ASSIGN_SIZE_COLUMN(month_trade_df)
    
    #stone type summery
    month_stone_trade_df = month_trade_df[month_trade_df['comodity'] == 'سنگ بریده']
    
    for stone_type in month_stone_trade_df['stone_type'].unique():
        stone_type_trade_df = month_stone_trade_df[month_stone_trade_df['stone_type']==stone_type]
        for stone_name in stone_type_trade_df['stone_name'].unique():
            stone_name_trade_df = stone_type_trade_df[stone_type_trade_df['stone_name']==stone_name]
            stone_meterage_trade = stone_name_trade_df['meterage'].sum()
            stone_pays_trade = currency_formart(int(stone_name_trade_df['total_pay'].sum()))
            
            avg_trade_price = currency_formart(int(stone_name_trade_df['total_pay'].sum()/stone_name_trade_df['meterage'].sum()))
            
            temp = [stone_type ,stone_name ,'', stone_meterage_trade , '',
                                            stone_pays_trade ,'' ,avg_trade_price]
            month_report_df.loc[len(month_report_df)] = temp
            
            for size in stone_name_trade_df['size'].unique():
                stone_name_size_trade_df = stone_name_trade_df[stone_name_trade_df['size'] == size]
                stone_name_size_meterage = stone_name_size_trade_df['meterage'].sum()
                stone_name_size_payed = currency_formart(int(stone_name_size_trade_df['total_pay'].sum()))
                
                stone_pays_details = {}
                stone_size_meterage_details = {}
            
                for trader in stone_name_size_trade_df['full_name'].unique():
                    stone_trader_trades = stone_name_size_trade_df[stone_name_size_trade_df['full_name']==trader]
                    stone_trader_meterage =stone_trader_trades['meterage'].sum()
                    stone_trader_pays = currency_formart(int(stone_trader_trades['total_pay'].sum()))
                    stone_pays_details[trader] = stone_trader_pays
                    stone_size_meterage_details[trader] = stone_trader_meterage
                    
                    avg_trade_price = currency_formart(int(stone_trader_trades['total_pay'].sum()/stone_trader_trades['meterage'].sum()))
                    
                temp = ['' ,'' ,size ,stone_name_size_meterage ,stone_size_meterage_details,
                        stone_name_size_payed ,stone_pays_details , avg_trade_price]
                month_report_df.loc[len(month_report_df)] = temp
                
                
    return month_report_df


def SUMMERIZE_MONTH_TRADES_DF2(month_buys_df,month_sells_df):
        
    month_report_df_columns = ['comodity', 'stone_type' ,'stone_name' ,'sizes' ,
                               'buy_meterage_total', 'buy_meterage_details', 'paid_total', 'paid_details','average_buy_price' ,'average_buy_price_details',
                               'sell_meterage_total', 'sell_meterage_details', 'earned_total', 'earned_details','average_sell_price' ,'average_sell_price_details' ]

    trader_type = ['seller_name','buyer_name']
    
    month_report_df = pandas.DataFrame(columns=month_report_df_columns)
    
    size_column= []
    
    month_buys_df = ASSIGN_SIZE_COLUMN(month_buys_df)
    month_sells_df = ASSIGN_SIZE_COLUMN(month_sells_df)
    
    month_total_trades_df = pandas.concat([month_buys_df , month_sells_df] , ignore_index=True)
    
    for comodity in month_total_trades_df['comodity'].unique():
        comodity_trades_df = month_total_trades_df[month_total_trades_df['comodity']==comodity].reset_index()
        
        # if comodity == 'سنگ بریده':
            
        for stone_type in month_total_trades_df['stone_type'].unique():
            stone_type_trades_df = month_total_trades_df[(month_total_trades_df['stone_type']==stone_type)]
            stone_type_buy_df = month_total_trades_df[(month_total_trades_df['stone_type']==stone_type) & (month_total_trades_df['seller_name'])]
            stone_type_sell_df = month_total_trades_df[(month_total_trades_df['stone_type']==stone_type) & (month_total_trades_df['buyer_name'])]
            
            for stone_name in stone_type_trades_df['stone_name'].unique():
                
                stone_name_trades_df = stone_type_trades_df[stone_type_trades_df['stone_name'] == stone_name]
                stone_name_buy_df = stone_type_buy_df[stone_type_buy_df['stone_name']==stone_name]
                stone_name_sell_df = stone_type_sell_df[stone_type_sell_df['stone_name']==stone_name]
                
                stone_meterage_buy = stone_name_buy_df['meterage'].sum() if stone_name_buy_df['meterage'].sum() > 0 else 0
                stone_meterage_sell = stone_name_sell_df['meterage'].sum() if stone_name_sell_df['meterage'].sum() > 0 else 0 
                
                stone_pays_buy = currency_formart(int(stone_name_buy_df['total_pay'].sum())) if stone_meterage_buy > 0 else 0
                stone_pays_sell = currency_formart(int(stone_name_sell_df['total_pay'].sum())) if stone_meterage_sell > 0 else 0
                
                avg_buy_price = currency_formart(int(stone_name_buy_df['total_pay'].sum()/stone_name_buy_df['meterage'].sum())) if stone_meterage_buy > 0 else 0
                avg_sell_price = currency_formart(int(stone_name_sell_df['total_pay'].sum()/stone_name_sell_df['meterage'].sum())) if stone_meterage_sell > 0 else 0
                
                    
                temp = [comodity ,stone_type ,stone_name ,'',
                        stone_meterage_buy,'', stone_pays_buy ,'' ,avg_buy_price,'',
                        stone_meterage_sell,'', stone_pays_sell ,'' ,avg_sell_price,'']
                
                month_report_df.loc[len(month_report_df)] = temp
                
                for size in stone_name_trades_df['size'].unique():
                    stone_size_trades_df = stone_name_trades_df[stone_name_trades_df['size']==size]
                    stone_size_buy_df = stone_name_buy_df[stone_name_buy_df['size'] == size]
                    stone_size_sell_df = stone_name_sell_df[stone_name_sell_df['size'] == size]
                    
                    stone_size_meterage_buy = stone_size_buy_df['meterage'].sum()
                    stone_size_meterage_sell = stone_size_sell_df['meterage'].sum()
                    
                    stone_size_payed_buy = currency_formart(int(stone_size_buy_df['total_pay'].sum()))
                    stone_size_payed_sell = currency_formart(int(stone_size_sell_df['total_pay'].sum()))
                    
                    stone_name_size_avg_buy_price = 0
                    stone_name_size_avg_sell_price = 0
                    
                    if not stone_size_buy_df['meterage'].empty:
                        stone_name_size_avg_buy_price = currency_formart(int(((stone_size_buy_df['meterage']*stone_size_buy_df['fee']).sum())/stone_size_meterage_buy))
                                            
                    if not stone_size_sell_df['meterage'].empty:
                        stone_name_size_avg_sell_price = currency_formart(int(((stone_size_sell_df['meterage']*stone_size_sell_df['fee']).sum())/stone_size_meterage_sell))
                        
                    stone_size_pays_details_buy = {}
                    stone_pays_details_sell = {}
                        
                    stone_size_meterage_details_buy = {}
                    stone_size_meterage_details_sell = {}
                    
                    stone_average_price_details_buy = {}
                    stone_average_price_details_sell = {}
                    
                    for seller in stone_size_buy_df['full_name'].unique():
                        stone_size_seller_buys = stone_size_buy_df[stone_size_buy_df['full_name']==seller]
                        stone_size_seller_meterage =stone_size_seller_buys['meterage'].sum()
                        stone_size_seller_pays = currency_formart(int(stone_size_seller_buys['total_pay'].sum()))
                        stone_size_pays_details_buy[seller] = stone_size_seller_pays
                        stone_size_meterage_details_buy[seller] = stone_size_seller_meterage
                        
                        avg_buy_price = currency_formart(int(stone_size_seller_buys['total_pay'].sum()/stone_size_seller_buys['meterage'].sum()))
                        stone_average_price_details_buy[seller] = avg_buy_price
                        
                    for buyer in stone_size_sell_df['full_name'].unique():
                        stone_buyer_trades = stone_size_sell_df[stone_size_sell_df['full_name']==buyer]
                        stone_buyer_meterage =stone_buyer_trades['meterage'].sum()
                        stone_buyer_pays = currency_formart(int(stone_buyer_trades['total_pay'].sum()))
                        stone_pays_details_sell[buyer] = stone_buyer_pays
                        stone_size_meterage_details_sell[buyer] = int(stone_buyer_meterage)
                        
                        avg_sell_price = currency_formart(int(stone_buyer_trades['total_pay'].sum()/stone_buyer_trades['meterage'].sum()))
                        stone_average_price_details_sell[buyer] = avg_sell_price
                        
                    temp = ['','' ,'' ,size ,
                            stone_size_meterage_buy ,stone_size_meterage_details_buy,
                            stone_size_payed_buy ,stone_size_pays_details_buy ,stone_name_size_avg_buy_price,stone_average_price_details_buy,
                            stone_size_meterage_sell ,stone_size_meterage_details_sell,
                            stone_size_payed_sell ,stone_pays_details_sell, stone_name_size_avg_sell_price ,stone_average_price_details_sell]
                    
                    month_report_df.loc[len(month_report_df)] = temp

        # else:
        #     for i in range(len(comodity_trades_df)):
        #         item = comodity_trades_df.loc[i]
        #         if not is_nan(item['seller_name']):
        #             temp = [item['comodity'], item['stone_type'] ,item['stone_name'] ,'',
        #                     item['meterage'], '', currency_formart(item['total_pay']), {item['full_name']:currency_formart(item['total_pay'])},'' ,'',
        #                     '', '', '', '','' ,'' ]
                
        #         elif not is_nan(item['buyer_name']):
        #             temp = [item['comodity'], item['stone_type'] ,item['stone_name'] ,'',
        #                     '', '', '', '','' ,'',
        #                     '', item['meterage'], currency_formart(item['total_pay']), {item['full_name']:currency_formart(item['total_pay'])},'' ,'' ]
                

                # month_report_df.loc[len(month_report_df)] = temp
    
    
    return month_report_df



def GET_MONTH_TRANSFERS_DF(month_folder):
    for file in os.listdir(month_folder):
        if file[:2].isdigit() and 'Payments' in file:
            month_transfers_list = file
    
    month_transfers_df = pandas.read_excel(pathlib.Path(str(month_folder) + '/' + month_transfers_list),header=0)
    month_transfers_df.dropna(subset=['transfer_type'] ,inplace=True)
    month_transfers_df['trader_full_name'] = month_transfers_df['trader_type']+ ' ' + month_transfers_df['trader_name']
    return month_transfers_df

def SUMMERIZE_MONTH_TRANSFERS_DF(month_transfers_df):
    month_transfers_report_df_columns = ['report','amount']
    month_transfers_report_df = pandas.DataFrame(columns=month_transfers_report_df_columns)
    
    total_pays= currency_formart(month_transfers_df['amount'].sum())
    temp = ['کل پرداخت' , total_pays]
    
    month_transfers_report_df.loc[len(month_transfers_report_df)] = temp

    for trader in month_transfers_df['trader_full_name'].unique():
        trader_payments_df = month_transfers_df[month_transfers_df['trader_full_name'] == trader]
        trader_total_payments = currency_formart(trader_payments_df['amount'].sum())
        
        temp = ['پرداخت به ' + trader , trader_total_payments]
        month_transfers_report_df.loc[len(month_transfers_report_df)] = temp
    
    return month_transfers_report_df

def CALCULATE_MONTH_STOCK_AVALABILITY(month_stone_buys_df, month_stone_sells_df):
    stock_df_columns = ['stone_type', 'stone_name' ,'stone_size', 'meterage' ,
                        'seller', 'average_buy_price' ,'total_paid_for' ,'buy_date']
    
    stock_df = pandas.DataFrame(columns=stock_df_columns)
    
    month_stone_buys_df = ASSIGN_SIZE_COLUMN(month_stone_buys_df)
    month_stone_sells_df = ASSIGN_SIZE_COLUMN(month_stone_sells_df)
    
    #stone type summery
    for stone_type in month_stone_buys_df['stone_type'].unique():
        stone_type_buy_df = month_stone_buys_df[month_stone_buys_df['stone_type']==stone_type]
        stone_type_sell_df = month_stone_sells_df[month_stone_sells_df['stone_type']==stone_type]
        
        for stone_name in stone_type_buy_df['stone_name'].unique():
            stone_name_buy_df = stone_type_buy_df[stone_type_buy_df['stone_name']==stone_name]
            stone_name_sell_df = stone_type_sell_df[stone_type_sell_df['stone_name']==stone_name]
            
            for size in stone_name_buy_df['size'].unique():
                stone_name_size_buy_df = stone_name_buy_df[stone_name_buy_df['size'] == size]
                stone_name_size_sell_df = stone_name_sell_df[stone_name_sell_df['size'] == size]
                
                stone_size_meterage_buy = stone_name_size_buy_df['meterage'].sum()
                stone_size_meterage_sell = 0
                if not stone_name_size_sell_df.empty:
                    stone_size_meterage_sell = stone_name_size_sell_df['meterage'].sum()
                
                stone_name_size_payed_buy = currency_formart(int(stone_name_size_buy_df['total_pay'].sum()))
                stone_size_payed_sell = 0
                if not stone_name_size_sell_df.empty:
                    stone_size_payed_sell = currency_formart(int(stone_name_size_sell_df['total_pay'].sum()))
                
                stone_size_stock=0
                if stone_size_meterage_buy > stone_size_meterage_sell:
                    stone_size_stock = stone_size_meterage_buy - stone_size_meterage_sell
                    
                if stone_size_stock > 0:
                    sellers = list(stone_name_size_buy_df['full_name'])
                    for seller in sellers:
                        seller_buy_df = stone_name_size_buy_df[stone_name_size_buy_df['full_name']==seller]
                        seller_meterage_buy = seller_buy_df['meterage']
                        
                    temp= [stone_type, stone_name ,size, stone_size_stock ,
                           'seller', 'average_buy_price' ,'total_paid_for' ,'buy_date']
        
def SUMMERIZE_PARTIES_ACCOUNTS(month_transfers_df,monthly_trades_report,month_stone_buys_df,month_stone_sells_df, save=True):
    end_month_parties_accounts_df = {}
    parties = list(month_transfers_df['trader_full_name'].unique())
    
    for party in month_stone_buys_df['full_name'].unique():
        if party not in parties:
            parties.append(party)
    
    for party in month_stone_sells_df['full_name'].unique():
        if party not in parties:
            parties.append(party)
    
    for party in parties:
        party_account_dic = {}
        buys_from_trader = month_stone_buys_df[month_stone_buys_df['full_name'] == party].reset_index()
        sells_to_trader = month_stone_sells_df[month_stone_sells_df['full_name'] == party].reset_index()
        trader_transactions = month_transfers_df[month_transfers_df['trader_full_name'] == party].reset_index()
        
        party_account_dic['buys'] = buys_from_trader.drop(columns=['full_name','size'])
        party_account_dic['sells'] = sells_to_trader.drop(columns=['full_name','size'])
        party_account_dic['transactions'] = trader_transactions.drop(columns=['trader_full_name'])
        
        party_account_dic['report'] = CALCULATE_PARTY_REPORT(party_account_dic)
        end_month_parties_accounts_df[party] = party_account_dic['report'][party_account_dic['report']['report']=='party_account'].reset_index().loc[0]['value']
        party_account_dic['report']=CHANGE_COLUMN_TO_CURRENCY_FORMAT(party_account_dic['report'],['value'])
        
        party_account_dic['buys'] = CHANGE_COLUMN_TO_CURRENCY_FORMAT(party_account_dic['buys'],['fee' , 'total_pay'])
        party_account_dic['sells'] = CHANGE_COLUMN_TO_CURRENCY_FORMAT(party_account_dic['sells'],['fee' , 'total_pay'])
        party_account_dic['transactions'] = CHANGE_COLUMN_TO_CURRENCY_FORMAT(party_account_dic['transactions'],['amount'])
        
        if save:
            party_account_save_path = month_folder + month_parties_report_folder_name + party + ".xlsx"
            # SAVE_DFDICT_TO_EXCELL_SHEETS(PERSIAN_DIC_DF(party_account_dic) , party_account_save_path)
            SAVE_DFDICT_TO_EXCELL_SHEETS(party_account_dic , party_account_save_path)
    
    return end_month_parties_accounts_df
        

def CALCULATE_PARTY_REPORT(party_account_dic):
    party_report_df_columns = ['report' , 'value']
    party_report_df = pandas.DataFrame(columns=party_report_df_columns)
    
    report ={}
    party_buys = party_account_dic['buys']
    party_sells = party_account_dic['sells']
    party_transactions = party_account_dic['transactions']
    
    report['buy_value_from_party'] = party_buys['total_pay'].sum()
    report['sell_value_to_party'] = party_sells['total_pay'].sum()
    
    report['paid_to_party'] = party_transactions[party_transactions['transfer_type']=='pay']['amount'].sum()
    report['recieved_from_party'] = party_transactions[party_transactions['transfer_type']=='receive']['amount'].sum()
    
    report['party_account'] = int(report['buy_value_from_party'] - report['paid_to_party'] + report['recieved_from_party'] - report['sell_value_to_party'])
    
    
    for key in report.keys():
        temp = [key , report[key]]
        party_report_df.loc[len(party_report_df)]=temp
    
    return party_report_df

def GET_TOTAL_OF_PARTY(party_account_dic):
    party_total_df_columns = ['action' ,'comodity', 'type', 'name' ,'size' ,'meter' ,'fee' ,'value' ,'date' ,'party_acount']
    
    party_total_df = pandas.DataFrame(columns=party_total_df_columns)
    
    party_buys = party_account_dic['buys']
    party_sells = party_account_dic['sells']
    party_transactions = party_account_dic['transactions']
    
    for i in range(len(party_buys)):
        item = party_buys.loc[i]
        temp = [item['trade_type'], item['comodity'] , item['stone_type'] ,item['stone_name'],
                item['size'] , str('%.2f' %item['meterage']),item['fee'],
                item['total_pay'],item['date'], '']
        party_total_df.loc[len(party_total_df)] = temp
        
    for i in range(len(party_sells)):
        item = party_sells.loc[i]
        temp = [item['trade_type'], item['comodity'] , item['stone_type'] ,item['stone_name'],
                item['size'] , str('%.2f' %item['meterage']),item['fee'],
                item['total_pay'],item['date'], '']
        party_total_df.loc[len(party_total_df)] = temp
        
    for i in range(len(party_transactions)):
        item = party_transactions.loc[i]
        temp = [item['transfer_type'], '' , '' , '','','','', item['amount'],
                item['send_date'], '']
        party_total_df.loc[len(party_total_df)] = temp
    
    
    party_total_df.sort_values(by=['date'], inplace=True, ignore_index=True)
    
    # party_account = 0
    item = party_total_df.loc[0]
    
    if item['action'] == 'sell' or item['action'] == 'پرداخت' or item['action'] == 'pay':
        party_total_df.at[0,'party_acount'] = -currency_formart(party_total_df.at[0,'value'], True)
    else:
        party_total_df.at[0,'party_acount'] = currency_formart(party_total_df.at[0,'value'] , True)

    for i in range(1,len(party_total_df)):
        item = party_total_df.loc[i]
        
        if item['action'] == 'buy' or item['action'] == 'دریافت' or item['action'] == 'receive':
            party_total_df.at[i,'party_acount'] = party_total_df.loc[i-1].get('party_acount',0) + currency_formart(party_total_df.loc[i].at['value'] ,True)
        
        elif item['action'] == 'sell' or item['action'] == 'پرداخت' or item['action'] == 'pay':
            party_total_df.at[i,'party_acount'] = party_total_df.loc[i-1].get('party_acount',0) - currency_formart(party_total_df.loc[i].at['value'] , True)
        
    
    CHANGE_COLUMN_TO_CURRENCY_FORMAT(party_total_df, ['party_acount'])
        
    return party_total_df        

    
def CALCULATE_MONTH_ACCOUNT_OF_PARTIES(month_payments_file):
    pass
  

def CALCULATE_END_MONTH_ACCOUNTS(end_month_parties_accounts_df, save=True):
    end_month_accounts_df_columns = ['party' , 'account']
    end_month_accounts_df = pandas.DataFrame(columns= end_month_accounts_df_columns)
    for party in list(end_month_parties_accounts_df.keys()):
        party_account = end_month_parties_accounts_df[party]
        
        end_month_accounts_df.loc[len(end_month_accounts_df)] = [party , party_account]
    
    end_month_accounts_df_columns = CHANGE_COLUMN_TO_CURRENCY_FORMAT(end_month_accounts_df ,['account'])
    if save:
        end_month_accounts_df_save_path =  month_folder + month_parties_report_folder_name + "01-حساب اخر ماه.xlsx"
        # SAVE_DFDICT_TO_EXCELL_SHEETS(PERSIAN_DIC_DF({'end_month_report':end_month_accounts_df}) , end_month_accounts_df_save_path)
        SAVE_DFDICT_TO_EXCELL_SHEETS({'end_month_report':end_month_accounts_df} , end_month_accounts_df_save_path)


def RECORD_PARTIES_ALL_ACCOUNT(monthes_folders_list):
    total_buys_df = GET_MONTH_TRADES_DF_DIC(monthes_folders_list[0],'buy')['sheet1']
    total_buys_df = ASSIGN_FULL_NAME_COLUMN(total_buys_df, 'seller_type', 'seller_name')
    total_buys_df = ASSIGN_SIZE_COLUMN(total_buys_df)
    
    total_sells_df = GET_MONTH_TRADES_DF_DIC(monthes_folders_list[0],'sell')['sheet1']
    total_sells_df = ASSIGN_FULL_NAME_COLUMN(total_sells_df, 'buyer_type', 'buyer_name')
    total_sells_df = ASSIGN_SIZE_COLUMN(total_sells_df)
    
    total_transfers_df = GET_MONTH_TRANSFERS_DF(monthes_folders_list[0])
    
    for i in range(1,len(monthes_folders_list)):
        temp_buy_df = GET_MONTH_TRADES_DF_DIC(monthes_folders_list[i],'buy')['sheet1']
        temp_buy_df = ASSIGN_FULL_NAME_COLUMN(temp_buy_df, 'seller_type', 'seller_name')
        temp_buy_df = ASSIGN_SIZE_COLUMN(temp_buy_df)    
        
        temp_sell_df = GET_MONTH_TRADES_DF_DIC(monthes_folders_list[i],'sell')['sheet1']
        temp_sell_df = ASSIGN_FULL_NAME_COLUMN(temp_sell_df, 'buyer_type', 'buyer_name')
        temp_sell_df = ASSIGN_SIZE_COLUMN(temp_sell_df)
        
        temp_transfers_df = GET_MONTH_TRANSFERS_DF(monthes_folders_list[i])
        total_transfers_df = pandas.concat([total_transfers_df , temp_transfers_df],axis=0,ignore_index=True)
        
        total_buys_df = pandas.concat([total_buys_df , temp_buy_df],axis=0,ignore_index=True)
        total_sells_df = pandas.concat([total_sells_df , temp_sell_df],axis=0,ignore_index=True)
        
    total_parties = list(total_buys_df['full_name'].unique())
    for item in list(total_sells_df['full_name'].unique()):
        if item not in total_parties:
            total_parties.append(item)
    
    for item in list(total_transfers_df['trader_full_name'].unique()):
        if item not in total_parties:
            total_parties.append(item)
            
    for party in total_parties:
        if 'احمدپور' in party:
            pass
        party_account_dic = {}
        
        buys_from_trader = total_buys_df[total_buys_df['full_name'] == party].reset_index()
        sells_to_trader = total_sells_df[total_sells_df['full_name'] == party].reset_index()
        trader_transactions = total_transfers_df[total_transfers_df['trader_full_name'] == party].reset_index()
        
        party_account_dic['buys'] = buys_from_trader.drop(columns=['full_name','index']).sort_values(by=['date'])
        party_account_dic['buys'].insert(loc=0,column='trade_type', value='buy')
        
        party_account_dic['sells'] = sells_to_trader.drop(columns=['full_name','index']).sort_values(by=['date'])
        party_account_dic['sells'].insert(loc=0,column='trade_type', value='sell')
        
        party_account_dic['transactions'] = trader_transactions.drop(columns=['trader_full_name','index']).sort_values(by=['send_date'])
        
        party_account_dic['report'] = CALCULATE_PARTY_REPORT(party_account_dic)
        # end_month_parties_accounts_df[party] = party_account_dic['report'][party_account_dic['report']['report']=='party_account'].reset_index().loc[0]['value']
        party_account_dic['report']=CHANGE_COLUMN_TO_CURRENCY_FORMAT(party_account_dic['report'],['value'])
        
        party_account_dic['buys'] = CHANGE_COLUMN_TO_CURRENCY_FORMAT(party_account_dic['buys'],['fee' , 'total_pay'])
        party_account_dic['sells'] = CHANGE_COLUMN_TO_CURRENCY_FORMAT(party_account_dic['sells'],['fee' , 'total_pay'])
        party_account_dic['transactions'] = CHANGE_COLUMN_TO_CURRENCY_FORMAT(party_account_dic['transactions'],['amount'])
        
        party_account_dic['total_report'] = GET_TOTAL_OF_PARTY(party_account_dic) 
        
        # PERSIAN_OF_DF(party_account_dic['buys'])
        
        party_account_value = currency_formart(party_account_dic['report'].loc[len(party_account_dic['report'])-1]['value'],True)
        if abs(party_account_value) <= 50:
            # SAVE_DFDICT_TO_EXCELL_SHEETS(PERSIAN_DIC_DF(party_account_dic) , directory + "//99-total_accounts//paid//"+party+".xlsx",True)
            SAVE_DFDICT_TO_EXCELL_SHEETS(party_account_dic , pathlib.Path(directory + '/' + '99-total_accounts' + '/' + 'paid' + '/' + party +".xlsx"),True)
        else:
            # SAVE_DFDICT_TO_EXCELL_SHEETS(PERSIAN_DIC_DF(party_account_dic) , directory + "//99-total_accounts//"+party+".xlsx",True)
            SAVE_DFDICT_TO_EXCELL_SHEETS(party_account_dic, pathlib.Path(directory + '/' + '99-total_accounts' + '/' +party+".xlsx") ,True)

    
            
             
#%% PARAMETERS

month_report_folder_name = variables.month_report_folder_name
month_parties_report_folder_name = variables.month_parties_report_folder_name


monthes_folders_list = GET_MONTHES_FOLDERS()

for month_folder in monthes_folders_list:
    month_name = month_folder.name
    month_name = month_name[3:]

    month_buys_df = GET_MONTH_TRADES_DF_DIC(month_folder,'buy')['sheet1']
    month_buys_df = ASSIGN_FULL_NAME_COLUMN(month_buys_df, 'seller_type', 'seller_name')
    
    month_buys_report_df = SUMMERIZE_MONTH_STONE_TRADES_DF(month_buys_df,'buy')
    
    
    month_sells_df = GET_MONTH_TRADES_DF_DIC(month_folder,'sell')['sheet1']
    month_sells_df = ASSIGN_FULL_NAME_COLUMN(month_sells_df, 'buyer_type', 'buyer_name')
    
    month_sells_report_df = SUMMERIZE_MONTH_STONE_TRADES_DF(month_sells_df,'sell')
    
    
    monthly_trades_report = SUMMERIZE_MONTH_TRADES_DF2(month_buys_df, month_sells_df)
    monthly_trades_report_save_path = pathlib.Path(str(month_folder) + '/' + month_report_folder_name + '/monthly_report.xlsx')
    monthly_trades_df_dic = {'monthly_report':monthly_trades_report} 
    # SAVE_DFDICT_TO_EXCELL_SHEETS(PERSIAN_DIC_DF(monthly_trades_df_dic) , monthly_trades_report_save_path)
    SAVE_DFDICT_TO_EXCELL_SHEETS(monthly_trades_df_dic , monthly_trades_report_save_path)
    
    
    month_transfers_df = GET_MONTH_TRANSFERS_DF(month_folder)
    month_transfers_report_df = SUMMERIZE_MONTH_TRANSFERS_DF(month_transfers_df)
    month_transfers_report_df_save_path = pathlib.Path(str(month_folder) + '/' + month_report_folder_name + '/monthly_payments.xlsx')
    # SAVE_DFDICT_TO_EXCELL_SHEETS(PERSIAN_DIC_DF({'monthly_payments':month_transfers_report_df}), month_transfers_report_df_save_path)
    SAVE_DFDICT_TO_EXCELL_SHEETS({'monthly_payments':month_transfers_report_df}, month_transfers_report_df_save_path)
    
    
    # end_month_parties_accounts_df = SUMMERIZE_PARTIES_ACCOUNTS(month_transfers_df, monthly_trades_report,ASSIGN_SIZE_COLUMN(month_buys_df),ASSIGN_SIZE_COLUMN(month_sells_df))
    # CALCULATE_END_MONTH_ACCOUNTS(end_month_parties_accounts_df)
    
    # CALCULATE_MONTH_STOCK_AVALABILITY(month_buys_df,month_sells_df)
    
    
RECORD_PARTIES_ALL_ACCOUNT(monthes_folders_list)



