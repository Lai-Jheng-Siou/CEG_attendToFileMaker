import requests
from datetime import datetime, timedelta
import os

from detectAbsence import detectAbsence
from Py_ODBC_FM import filemaker_odbc_connection

from tools.custTool import getConfig


def main():
    date = datetime.today() - timedelta(days=1)
    year = date.strftime('%Y')
    month = date.strftime('%m ')
    day = date.strftime('%d')
    timestamp = int(datetime.timestamp(date))

    config = getConfig()

    #url
    url = config['http']['url_main']
    searchUrl = config['http']['url_search'] + f'q=(workno_date~equals~{year}%2F{month}%2F{day}~date2)'
    exportUrl = config['http']['url_export']

    #account & password
    acc = config['http']['account']
    pasd = config['http']['password']

    payload_v1 = {
        'username': acc,
        'password': pasd,
        'btnSubmit': 'Login'
    }

    session = requests.session()
    
    r = session.post(url, data = payload_v1)

    r = session.get(searchUrl)

    session_code = session.cookies.get_dict()

    payload_v2 = {
        "type": "excel2007",
        "records": "all",
        "txtformatting": "formatted",
        "page": "export",
    }

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
        "Referer": "https://ip.ceg.com.tw/attend_export.php",
        "Cookie": f"_ga_DM0K38FZR9=GS1.1.{timestamp}.1.1.{timestamp + 100}.0.0.0; _ga_HTDVDM2P6L=GS1.1.{timestamp}.1.1.{timestamp + 100}.0.0.0; _ga=GA1.3.698881609.{timestamp}; s1665029584={session_code['s1665029584']}; mediaType=1",
        "Content-Type": "application/x-www-form-urlencoded"
    }

    r = session.post(exportUrl, data = payload_v2, headers = headers)
    st = f'status_code: {r.status_code}, '
    if r.status_code == 200:
        # 將二進制數據保存為 Excel 文件
        with open("attend.xlsx", "wb") as f:
            f.write(r.content)
        st += '檔案下載成功\n'
    else:
        st += '檔案下載失敗\n'

    return st


if __name__ == "__main__":
    result = main()

    path = os.path.abspath(os.getcwd())
    xlsxPath = path + "\\attend.xlsx"

    st = detectAbsence(xlsxPath)

    filemaker_odbc_connection(xlsxPath)

    desktop = os.path.join(os.path.expanduser("~"), 'Desktop') + "\\attend.txt"


    with open(desktop, 'a') as txt:
        txt.write(result)
        txt.write(st)
        try:
            os.remove(xlsxPath)  # 刪除指定位子的檔案
            txt.write(f"檔案 {xlsxPath} 刪除成功\n")
        except FileNotFoundError:
            txt.write(f"檔案 {xlsxPath} 不存在\n")
        except PermissionError:
            txt.write(f"無法刪除檔案 {xlsxPath}，沒有足夠的權限\n")
        except Exception as e:
            txt.write(f"刪除檔案 {xlsxPath} 時發生錯誤：{e}\n")