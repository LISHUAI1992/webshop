import xlrd
import pymysql

book = xlrd.open_workbook(r'C:\test.xls')
sheet = book.sheet_by_name("Sheet1")
#建立一个MySQL连接
database = pymysql.connect (host="132.232.101.227", user = "myuser", passwd = "Hik19920623#123", db = "shop")
# 获得游标对象, 用于逐行遍历数据库数据
cursor = database.cursor()

# 创建一个for循环迭代读取xls文件每行数据的, 从第二行开始是要跳过标题
for r in range(1, sheet.nrows):
      order_indexcode = sheet.cell(r,0).value
      title_txt       = sheet.cell(r,1).value
      price           = sheet.cell(r,2).value
      number          = sheet.cell(r,3).value
      othersys_indexcode       = sheet.cell(r,4).value
      attribute       = sheet.cell(r,5).value
      package_info    = sheet.cell(r,6).value
      remarks         = sheet.cell(r,7).value
      status          = sheet.cell(r,8).value
      merchant_code   = sheet.cell(r,9).value

      #values = (order_indexcode, title_txt, price, number, othersys_indexcode, attribute, package_info, remarks, status, merchant_code)
      sql = "INSERT INTO commodity (order_indexcode, title_txt, price, number, othersys_indexcode, attribute, package_info, remarks, status, merchant_code)"\
      "VALUES ('%s', '%s', %d, %d, '%s', '%s', '%s', '%s', '%s', '%s')"%\
      (order_indexcode, title_txt, price, number, othersys_indexcode, attribute, package_info, remarks, status, merchant_code)
      print(sql)

      # 执行sql语句
      cursor.execute(sql)

# 关闭游标
cursor.close()

# 提交
database.commit()

# 关闭数据库连接
database.close()
