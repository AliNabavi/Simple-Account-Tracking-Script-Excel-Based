# -*- coding: utf-8 -*-
"""
Created on Wed Sep 27 21:11:38 2023

@author: ali-nbv
"""

save_as_persian = False


# save pathes
if save_as_persian:
    month_report_folder_name = 'گزارش-ماهانه' + '//'
    month_parties_report_folder_name = month_report_folder_name +'حساب طرفین' + '//'
else:
    month_report_folder_name = 'monthly_report' + '//'
    month_parties_report_folder_name = month_report_folder_name +'parties_account' + '//'
    


def GET_PERSIAN_NAME(eng_name):
    colomn_names_buys = {'buys':'خرید', 'sells':'فروش','transactions':'انتقال پول','total_report':'گزارش کلی حساب', 'trade_type':'معامله', 'comodity':'کالا' ,'stone_type':'نوع سنگ', 'stone_name':'نام سنگ','thickness':'ضخامت',
                    'length':'طول','width':'عرض','count':'تعداد','meterage':'متراژ', 'fee':'قیمت واحد','seller_type':'نوع فروشنده','seller_name':'نام فروشنده','total_pay':'قیمت کل','date':'تاریخ','of_buy_factor':'شماره فاکتور','size':'سایز','buy_factor_number':'شماره فاکتور'}
    colomn_names_sells = {'buyer_type':'نوع خریدار','buyer_name':'نام خریدار','of_buy_factor':'شماره فاکتوذ فروش'}
    colomn_names_transactions = {'transfer_type':'نوع انتقال','trader_name':'نام','trader_type':'نوع','amount':'مقدار','for':'بابت','send_date':'تاریخ ارسال','pay_date':'تاریخ سررسید' ,'payment_factor_number':'شماره فاکتور پرداخت'}
    colomn_names_report = {'report':'گزارش','value':'ارزش'}
    colomn_names_total = {'action':'عملیات','type':'نوع','name':'نام','meter':'متراژ','party_account':'مانده حساب'}
    value_names_report = {'buy_value_from_party':'خرید از طرف','sell_value_to_party':'فروش به طرف','paid_to_party':'مبلغ پرداختی به طرف','recieved_from_party':'مبلغ دریافتی از طرف','party_acount':'مانده حساب'}
    value_names_total = {'buy':'خرید','sell':'فروش'}
    columns_names_month_report = {'sizes':'سایز','buy_meterage_total':'متراژ خرید','buy_meterage_details':'ریز خرید','paid_total':'پرداخت شده','paid_details':'ریز پرداخت ها',
                                  'average_buy_price':'میانگین فی خرید','average_buy_price_details':'ریز میانگین خرید','sell_meterage_total':'متراژ فروخته شده','sell_meterage_details':'ریز متراژ فروش',
                                  'earned_total':'دریافتی از فروش','earned_details':'ریز دریافتی','average_sell_price':'میانگین فی فروش','average_sell_price_details':'ریز میانگین فروش'}
    if eng_name in columns_names_month_report:
        return columns_names_month_report[eng_name]
    
    if eng_name in colomn_names_buys:
        return colomn_names_buys[eng_name]
    
    if eng_name in colomn_names_sells:
        return colomn_names_sells[eng_name]
    
    if eng_name in colomn_names_transactions:
        return colomn_names_transactions[eng_name]
    
    if eng_name in colomn_names_report:
        return colomn_names_report[eng_name]
    
    if eng_name in colomn_names_total:
        return colomn_names_total[eng_name]
    
    if eng_name in value_names_report:
        return value_names_report[eng_name]
    
    if eng_name in value_names_total:
        return value_names_total[eng_name]
    
    return eng_name
     
    

