import pandas as pd


def cleanCRNs(excel_df: pd.DataFrame) -> pd.DataFrame:
    crnList = []

    for index, row in excel_df.iterrows():
        crnList.append(((str)(row["사업자등록번호"])).strip())

    excel_df["사업자등록번호"] = crnList
    return excel_df


excel1 = cleanCRNs(pd.read_excel('./assets/input1.xlsx'))
excel2 = cleanCRNs(pd.read_excel('./assets/input2.xlsx'))

nameList = []
addressList = []
postcodeList = []

for crn in excel1["사업자등록번호"]:
    targetRow = excel2.loc[excel2["사업자등록번호"] == crn]

    if targetRow.empty:
        nameList.append("일치하는 사업자 없음")
        addressList.append("일치하는 사업자 없음")
    else:
        nameList.append(targetRow.iloc[0]["대표자이름"].strip())
        targetAddress = ((str) (targetRow.iloc[0]["대표자주소"])).strip()
        targetPostcode = ((str) (targetRow.iloc[0]["대표자우편번호"])).strip()

        if targetAddress == "nan" or targetPostcode == "null":
            addressList.append((str) (targetRow.iloc[0]["소재지주소"]).strip())
            postcodeList.append("확인필요")
        else:
            while (len(targetPostcode) < 5):
                targetPostcode = "0" + targetPostcode

            addressList.append(targetAddress)
            postcodeList.append(targetPostcode)

excel1.insert(len(excel1.columns), "대표자이름", nameList, True)
excel1.insert(len(excel1.columns), "대표자주소", addressList, True)
excel1.insert(len(excel1.columns), "대표자우편번호", postcodeList, True)

excel1.to_excel('./assets/output.xlsx')
