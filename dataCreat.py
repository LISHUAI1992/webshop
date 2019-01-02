import xlrd
import pymysql
import traceback
import time
from datetime import datetime
from xlrd import xldate_as_tuple

def is_exit_order(cursor, indexcode):
    # SQL 查询语句
    sql = "SELECT * FROM `order` \
           WHERE order_indexcode = %s" % (indexcode)
    exit = 0;
    try:
        # 执行SQL语句
        cursor.execute(sql)
        # 获取所有记录列表
        results = cursor.fetchall()
        for row in results:
            exit = 1
    except:
        print
        "Error: unable to fecth data"
    return exit;

book = xlrd.open_workbook(r'C:\开发相关\数据\测试数据\订单1201.xls')
sheet = book.sheet_by_name(r"订单1202")

# 建立一个MySQL连接

database = pymysql.connect(host="132.232.101.227", user = "myuser", passwd = "Hik19920623#123", db = "shop")

# 获得游标对象, 用于逐行遍历数据库数据
cursor = database.cursor()



# sheet.nrows
# 创建一个for循环迭代读取xls文件每行数据的,	从第二行开始是要跳过标题
for r in range(1, sheet.nrows):
    list = []
    for l in range(0, 58):
        if ( l == 19 ) or (l == 20) or (l == 53):
            date = datetime(*xldate_as_tuple(sheet.cell(r, l).value, 0))
            temptime = date.strftime('%Y-%m-%d %H:%M:%S')
            list.append(temptime)
        else:
            list.append(sheet.cell(r, l).value)

    order_indexcode = sheet.cell(r, 0).value
    buyer_name = sheet.cell(r, 1).value
    b_pay_account = sheet.cell(r, 2).value
    b_pay_code = sheet.cell(r, 3).value
    b_pay_details = sheet.cell(r, 4).value
    b_pay_monay = sheet.cell(r, 5).value
    b_pay_postage = sheet.cell(r, 6).value
    b_pay_integral = sheet.cell(r, 7).value
    b_total_monay = sheet.cell(r, 8).value
    b_rebates_integral = sheet.cell(r, 9).value
    b_realpay_monay = sheet.cell(r, 10).value
    b_realpay_integral = sheet.cell(r, 11).value
    b_order_state = sheet.cell(r, 12).value
    b_note = sheet.cell(r, 13).value
    b_recipient_name = sheet.cell(r, 14).value
    b_recipient_adress = sheet.cell(r, 15).value
    o_transport_type = sheet.cell(r, 16).value
    b_recipient_phone = sheet.cell(r, 17).value
    b_recipient_mphone = sheet.cell(r, 18).value
    b_order_createtime = sheet.cell(r, 19).value
    date = datetime(*xldate_as_tuple(b_order_createtime, 0))
    b_order_createtime = date.strftime('%Y-%m-%d %H:%M:%S')

    b_order_paytime = sheet.cell(r, 20).value
    date = datetime(*xldate_as_tuple(b_order_paytime, 0))
    b_order_paytime = date.strftime('%Y-%m-%d %H:%M:%S')

    o_commodity_note = sheet.cell(r, 21).value
    o_commodity_type = sheet.cell(r, 22).value
    o_logistic_code = sheet.cell(r, 23).value
    o_logistic_company = sheet.cell(r, 24).value
    o_order_note = sheet.cell(r, 25).value
    o_commodity_count = sheet.cell(r, 26).value
    o_shop_id = sheet.cell(r, 27).value
    o_shop_name = sheet.cell(r, 28).value
    o_order_closereson = sheet.cell(r, 29).value
    s_seller_fee = sheet.cell(r, 30).value
    b_buyer_fee = sheet.cell(r, 31).value
    o_invoice_info = sheet.cell(r, 32).value
    o_phon_order = sheet.cell(r, 33).value
    o_phaseorder_info = sheet.cell(r, 34).value
    o_privilegeorder_id = sheet.cell(r, 35).value
    o_contract_pic = sheet.cell(r, 36).value
    o_order_receipts = sheet.cell(r, 37).value
    o_order_paid = sheet.cell(r, 38).value
    o_deposit_rank = sheet.cell(r, 39).value
    o_modified_sku = sheet.cell(r, 40).value
    o_modified_adress = sheet.cell(r, 41).value
    o_abnormal_info = sheet.cell(r, 42).value
    o_tmall_voucher = sheet.cell(r, 43).value
    o_jifenbao_voucher = sheet.cell(r, 44).value
    o_o2o_trading = sheet.cell(r, 45).value
    o_trading_type = sheet.cell(r, 46).value
    o_retailshop_name = sheet.cell(r, 47).value
    o_retailshop_id = sheet.cell(r, 48).value
    o_retaildelivery_name = sheet.cell(r, 49).value
    o_retaildelivery_id = sheet.cell(r, 50).value
    o_refund_account = sheet.cell(r, 51).value
    o_appointment_shop = sheet.cell(r, 52).value

    b_order_confirmtime = sheet.cell(r, 53).value
    date = datetime(*xldate_as_tuple(b_order_confirmtime, 0))
    b_order_confirmtime = date.strftime('%Y-%m-%d %H:%M:%S')

    b_pay_confirmaccount = sheet.cell(r, 54).value
    o_buyer_envelope = sheet.cell(r, 55).value
    o_mainorder_indexcode = sheet.cell(r, 56).value
    o_ext1_info = sheet.cell(r, 57).value
    o_ext2_info = sheet.cell(r, 58).value

    nowTime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    # nowTime = time.time().strftime('%Y-%m-%d %H:%M:%S')
    sql = ""

    if is_exit_order(cursor, order_indexcode) == 0:
        sql = "INSERT INTO `order` (order_indexcode, buyer_name, b_pay_account, b_pay_code, b_pay_details, b_pay_monay, b_pay_postage, b_pay_integral, b_total_monay, b_rebates_integral," \
          "b_realpay_monay, b_realpay_integral, b_order_state, b_note, b_recipient_name, b_recipient_adress, o_transport_type, b_recipient_phone, b_recipient_mphone, b_order_createtime, b_order_paytime," \
          "o_commodity_note, o_commodity_type, o_logistic_code, o_logistic_company, o_order_note, o_commodity_count, o_shop_id, o_shop_name, o_order_closereson, s_seller_fee, b_buyer_fee," \
          "o_invoice_info, o_phon_order, o_phaseorder_info, o_privilegeorder_id, o_contract_pic, o_order_receipts, o_order_paid, o_deposit_rank, o_modified_sku, o_modified_adress, o_abnormal_info," \
          "o_tmall_voucher, o_jifenbao_voucher, o_o2o_trading, o_trading_type, o_retailshop_name, o_retailshop_id, o_retaildelivery_name, o_retaildelivery_id, o_refund_account," \
          "o_appointment_shop, b_order_confirmtime, b_pay_confirmaccount, o_buyer_envelope, o_mainorder_indexcode, o_ext1_info, o_ext2_info, createtime)" \
          "VALUES ('%s', '%s', '%s', '%s', '%s', %s, %s, %s, %s, %s, %s, %s, '%s', '%s', '%s', '%s', '%s'," \
          "'%s', '%s', '%s','%s', '%s', '%s', '%s', '%s', '%s', '%s', %s, '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s'," \
          "'%s', '%s', '%s', '%s', '%s','%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s','%s', %s, '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % \
          (order_indexcode, buyer_name, b_pay_account, b_pay_code, b_pay_details, b_pay_monay, b_pay_postage,
           b_pay_integral, b_total_monay, b_rebates_integral, b_realpay_monay,
           b_realpay_integral, b_order_state, b_note, b_recipient_name, b_recipient_adress, o_transport_type,
           b_recipient_phone, b_recipient_mphone, b_order_createtime, b_order_paytime, o_commodity_note,
           o_commodity_type, o_logistic_code, o_logistic_company, o_order_note, o_commodity_count, o_shop_id,
           o_shop_name, o_order_closereson, s_seller_fee, b_buyer_fee, o_invoice_info,
           o_phon_order, o_phaseorder_info, o_privilegeorder_id, o_contract_pic, o_order_receipts, o_order_paid,
           o_deposit_rank, o_modified_sku, o_modified_adress, o_abnormal_info,
           o_tmall_voucher, o_jifenbao_voucher, o_o2o_trading, o_trading_type, o_retailshop_name, o_retailshop_id,
           o_retaildelivery_name, o_retaildelivery_id, o_refund_account,
           o_appointment_shop, b_order_confirmtime, b_pay_confirmaccount, o_buyer_envelope, o_mainorder_indexcode,
           o_ext1_info, o_ext2_info, nowTime)


    else:
        sql = "UPDATE `order`SET `buyer_name` = '%s',`b_pay_account` = '%s',`b_pay_code` = '%s'," \
              "`b_pay_details` = '%s',`b_pay_monay` = %s,`b_pay_postage` = %s," \
              "`b_pay_integral` = %s,`b_total_monay` = %s,`b_rebates_integral` = %s,`b_realpay_monay` = %s,`b_realpay_integral` = %s," \
              "`b_order_state` = '%s',`b_note` = '%s',`b_recipient_name` = '%s',`b_recipient_adress` = '%s'," \
              "`o_transport_type` = '%s',`b_recipient_phone` = '%s',`b_recipient_mphone` = '%s',`b_order_createtime` = '%s',`b_order_paytime` = '%s'," \
              "`o_commodity_note` = '%s',`o_commodity_type` = '%s',`o_logistic_code` = '%s',`o_logistic_company` = '%s',`o_order_note` = '%s'," \
              "`o_commodity_count` = %s,`o_shop_id` = '%s',`o_shop_name` = '%s',`o_order_closereson` = '%s',`s_seller_fee` = '%s',`b_buyer_fee` = '%s'," \
              "`o_invoice_info` = '%s',`o_phon_order` = '%s',`o_phaseorder_info` = '%s',`o_privilegeorder_id` = '%s',`o_contract_pic` = '%s'," \
              "`o_order_receipts` = '%s',`o_order_paid` = '%s',`o_deposit_rank` = '%s',`o_modified_sku` = '%s',`o_modified_adress` = '%s'," \
              "`o_abnormal_info` = '%s',`o_tmall_voucher` = '%s',`o_jifenbao_voucher` = '%s',`o_o2o_trading` = '%s',`o_trading_type` = '%s'," \
              "`o_retailshop_name` = '%s',`o_retailshop_id` = '%s',`o_retaildelivery_name` = '%s',`o_retaildelivery_id` = '%s',`o_refund_account` = '%s'," \
              "`o_appointment_shop` = %s,`b_order_confirmtime` = '%s',`b_pay_confirmaccount` = '%s',`o_buyer_envelope` = '%s'," \
              "`o_mainorder_indexcode` = '%s',`o_ext1_info` = '%s',`o_ext2_info` = '%s',`updatatime` = '%s' " \
              "WHERE" \
              "`order_indexcode` = Cast( '%s' AS BINARY ( 18 ) );" % \
              (buyer_name, b_pay_account, b_pay_code, b_pay_details, b_pay_monay, b_pay_postage,
               b_pay_integral, b_total_monay, b_rebates_integral, b_realpay_monay,
               b_realpay_integral, b_order_state, b_note, b_recipient_name, b_recipient_adress, o_transport_type,
               b_recipient_phone, b_recipient_mphone, b_order_createtime, b_order_paytime, o_commodity_note,
               o_commodity_type, o_logistic_code, o_logistic_company, o_order_note, o_commodity_count, o_shop_id,
               o_shop_name, o_order_closereson, s_seller_fee, b_buyer_fee, o_invoice_info,
               o_phon_order, o_phaseorder_info, o_privilegeorder_id, o_contract_pic, o_order_receipts, o_order_paid,
               o_deposit_rank, o_modified_sku, o_modified_adress, o_abnormal_info,
               o_tmall_voucher, o_jifenbao_voucher, o_o2o_trading, o_trading_type, o_retailshop_name, o_retailshop_id,
               o_retaildelivery_name, o_retaildelivery_id, o_refund_account,
               o_appointment_shop, b_order_confirmtime, b_pay_confirmaccount, o_buyer_envelope, o_mainorder_indexcode,
               o_ext1_info, o_ext2_info, nowTime, order_indexcode)
    print(sql)

    #	执行sql语句
    cursor.execute(sql)

# 关闭游标
cursor.close()

#	提交
database.commit()

#	关闭数据库连接
database.close()
