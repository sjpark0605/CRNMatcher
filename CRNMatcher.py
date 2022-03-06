import pandas as pd


def cleanCRNs(excel_df: pd.DataFrame):
    crnList = []

    for index, row in excel_df.iterrows():
        crnList.append(row["사업자등록번호"].strip())

    excel_df["사업자등록번호"] = crnList
    return excel_df


excel1 = cleanCRNs(pd.read_excel('./assets/input1.xlsx'))
excel2 = cleanCRNs(pd.read_excel('./assets/input2.xlsx'))

nameList = []
addressList = []

for crn in excel1["사업자등록번호"]:
    targetRow = excel2.loc[excel2["사업자등록번호"] == crn]

    if targetRow.empty:
        nameList.append("일치하는 사업자 없음")
        addressList.append("일치하는 사업자 없음")
    else:
        nameList.append(targetRow.iloc[0]["대표자명"].strip())
        addressList.append(targetRow.iloc[0]["주소"].strip())

excel1.insert(len(excel1.columns), "대표자명", nameList, True)
excel1.insert(len(excel1.columns), "주소", addressList, True)

excel1.to_excel('./assets/output.xlsx')
