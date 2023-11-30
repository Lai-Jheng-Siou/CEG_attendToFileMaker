import asyncio
from pyppeteer import launch
from datetime import datetime
from time import sleep
import os

from detectAbsence import detectAbsence
from Py_ODBC_FM import filemaker_odbc_connection

date = datetime.now()
lastDay = "%s/%s/%s"%(date.year, date.month, date.day - 1)
async def main():
    # 啟動瀏覽器
    browser = await launch(
        executablePath='C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe', #Chrome瀏覽器程式位置
        headless=False,
        )

    # 創建新的頁面
    page = await browser.newPage()

    # 前往目標網頁
    await page.goto('https://ip.ceg.com.tw/attend_search.php')

    await page.type('input[name="username"]', 'admin')
    await page.type('input[name="password"]', 'CegSystem@80017427')
    await page.click('#submitLogin1')

    await asyncio.sleep(1)

    await page.type('input[name="value_workno_date_1"]', lastDay)
    await page.click('#searchButton1')

    await asyncio.sleep(1)

    await page.goto('https://ip.ceg.com.tw/attend_export.php')
    await page.click('#saveButton1')

    await asyncio.sleep(5)

    # 關閉瀏覽器
    await browser.close()

if __name__ == "__main__":
    # asyncio.run(main())

    xlsxPath = r"C:\Users\CEG\Desktop\attend.xlsx"
    st = detectAbsence(xlsxPath)

    sleep(1)

    filemaker_odbc_connection(xlsxPath)

    sleep(1)

    with open(r"C:\Users\CEG\Desktop\attend.txt", 'a') as txt:
        txt.write(st)
        try:
            os.remove(xlsxPath)  # 刪除指定位子的檔案
            txt.write(f"檔案 {xlsxPath} 刪除成功\n")
        except FileNotFoundError:
            txt.write(f"檔案 {xlsxPath} 不存在\n")
        except PermissionError:
            txt.write(f"無法刪除檔案 {xlsxPath}，可能因為沒有足夠的權限\n")
        except Exception as e:
            txt.write(f"刪除檔案 {xlsxPath} 時發生錯誤：{e}\n")