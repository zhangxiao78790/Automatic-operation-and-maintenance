import telnetlib  # 导入telnetlib模块，用于远程登录OLT设备
import time      # 导入时间模块，用于获取当前时间戳和程序休眠
import xlrd as xl   # 导入xlrd模块，用于读取excel文件的数据


# 打开excel表，获取sheet
def open_excel(file_xls,sheet_index):
    """
    此函数用于获取excel表格以及sheet对象
    :param file_xls : excel表格文件 
    :param sheet_index : sheet索引
    :return table_sheet_ip : excel表格的ip sheet
    :return table_sheet_cmd : excel表格的cmd sheet
    """
    try:
        open_file = xl.open_workbook(file_xls) # 打开文件
        table_sheet_ip = open_file.sheet_by_index(0) # 获取文件中的ip地址的sheet
        table_sheet_cmd = open_file.sheet_by_index(sheet_index) # 获取文件中的配置的sheet
        return table_sheet_ip,table_sheet_cmd
    except FileNotFoundError:
        print("该文件在当前路径下找不到")
    except IndexError:
        print("输入的sheet索引超出范围")
     
    
# 获取OLT设备管理IP地址以及设备配置命令
def get_host_ip_cmd(file_xls,sheet_index):
    """
    此函数用于获取OLT设备管理IP地址
    :param file_xls : excel表格文件 
    :param sheet_index : sheet索引
    :return host_ip : OLT设备管理IP地址
    :return cmd : OLT设备的配置命令
    """
    host_ip = []   # OLT设备管理IP地址
    cmd = []    # OLT设备的配置命令
    table_sheet_ip,table_sheet_cmd = open_excel(file_xls,sheet_index) # 调用open_excel(file_xls)函数的ip sheet与cmd sheet
    try:
        for row_num in range (table_sheet_ip.nrows):    
            row_value = table_sheet_ip.row_values(row_num)
            host_ip.append(row_value[0])
        for row_num in range(table_sheet_cmd.nrows):
            row_value = table_sheet_cmd.row_values(row_num)
            cmd.append(row_value[0])
        return host_ip,cmd
    except Exception as result:
        print("出现异常错误:%s"%result)
                
    
# OLT设备登录
def login():
    """
    此函数用于登录OLT设备，输入账号密码
    """
    username = bytes(input("请输入账号:") +"\n",encoding='utf-8')
    password = bytes(input("请输入密码:") +"\n",encoding='utf-8')
    tn.read_until(b"Username:")
    tn.write(username)
    tn.read_until(b"Password:")
    tn.write(password)
    

# 对OLT设备进行配置的函数
def config():
    """
    此函数用于对OLT设备进行配置
    """
    login()
    for command in cmd:   
        tn.write(bytes(str(command)+"\n",encoding='utf-8'))
        time.sleep(2)
    time.sleep(1)
    out_result = tn.read_very_eager()
    out_result_str = out_result.decode()
    print(out_result_str)
    print(host + "——我已经做完了，请检查上述配置内容")
    tn.close()
    
# 主程序入口
if __name__ == '__main__':
    file_xls = "OLT_IP.xlsx"
    while True: 
        try:
            sheet_index = int(input("请输入整数型sheet索引:"))
            break
        except ValueError:
            print("请输入整数型")
    host_ip,cmd = get_host_ip_cmd(file_xls,sheet_index)  # 获取OLT管理ip地址以及配置命令的列表
    for host in host_ip:
        tn = telnetlib.Telnet(host,port=23,timeout=5)
        config()
    while True:
        print("*********************************************************************************************")
        print("*********************************************************************************************")
        over = input("请输入end结束本次操作:")
        if over == "end":
            break
        