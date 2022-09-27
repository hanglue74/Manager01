
import xlwings as xw

wb = xw.Book(r'F:/NICE/00.개인별요청자료/00. 구매업무/00. 2022년 KPI/재고자산 관리업무 협업/전산_자산관리_20220721.xlsm')
app = wb.app

macro_order = app.macro('순번')

macro_order()

# Second 파일에 대한 수정을 진행 중

## 첫번째 브랜치에 대한 수정 작업을 진행중...

# Second 파일에 대한 수정을 진행 중