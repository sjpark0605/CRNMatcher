from queue import Empty
import pandas as pd

excel1 = pd.read_excel('./assets/input1.xlsx')

nameList = []
addressList = []

excel2 = pd.read_excel('./assets/input2.xlsx')

for crn in excel1["사업자등록번호"]:
    key = (str)(crn.strip())
    targetRow = excel2.loc[excel2["사업자등록번호"] == key]

    if targetRow.empty:
        nameList.append("일치하는 사업자 없음")
        addressList.append("일치하는 사업자 없음")
    else:
        nameList.append(targetRow.iloc[0]["대표자명"])
        addressList.append(targetRow.iloc[0]["주소"])

excel1.insert(len(excel1.columns), "대표자명", nameList, True)
excel1.insert(len(excel1.columns), "주소", addressList, True)
