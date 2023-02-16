import pandas as pd
import time
import os
import gc
import re
import numpy as np
import openpyxl as xl
from openpyxl.styles import Font, Alignment
from tqdm import tqdm

#cashShopIdIndexList = cashShopIdList.index

cashShopDir = "./CL_CashShop"
if not os.path.isdir(cashShopDir) :
    os.mkdir(cashShopDir)

xlFileName = f"./CL_CashShop/result_{time.strftime('%y%m%d_%H%M%S')}.xlsx"

class Sales():
    pkgID = ""
    pkgName = ""
    category = ""
    desc = ""
    info = ""
    price = ""
    bonus = ""
    itemList0 = ""
    itemList1 = ""
    itemList2 = ""
    limit = ""
    server = ""
    
class Item():

    def __init__(self) :
        self.name = ""
        self.id = ""
        self.innerItemList = ""


def extract_data_cashshop(fileName):

    target = pd.read_csv(fileName)
    # fileName = "유료상점.xlsx"
    # target = pd.read_excel(fileName,sheet_name = '유료상점',engine='openpyxl')
    #target["CashShop ID"] = target["CashShop ID"].replace(n,0)

    target = target.replace('-',np.nan)
    cashShopIdList = target.drop_duplicates(subset='CashShopID')["CashShopID"]
    cashShopIdList = cashShopIdList.dropna(axis=0)
    cashShopIdIndexList = cashShopIdList.index

    totalCount = len(cashShopIdIndexList)
    #print(cashShopIdList.astype(int))
    #print(f'추가 상품 개수 : {totalCount}')

    gachaItemIndexList = target[["ItemID1","ItemID2"]].dropna(axis=0).index
    #cashShopIdList = cashShopIdList.dropna(axis=0)
    #print(gachaItemIndexList)
    salesList = [Sales] 
    #salesList : list[Sales]
    salesList.clear()
    print("데이터 추출 중...")
    tqdmCount2 =0
    for j in tqdm(range(0,totalCount)):
        tqdmCount2 +=1
        #print(cashShopIdIndexList[j], j+1)

        if (j+1) >= len(cashShopIdIndexList) :
            tempDf = target[cashShopIdIndexList[j]:]
        else :
            tempDf = target[cashShopIdIndexList[j]:cashShopIdIndexList[j+1]]
        tempDf = tempDf.reset_index()
        #for i in range(0,len(cashShopIdIndexList)):
        for i in range(0,1):
            a = Sales()
            a.pkgID = int(tempDf.loc[0,"CashShopID"])
            a.pkgName = tempDf.loc[0,"PkgName"] #+ "[귀속]"
            a.category = tempDf.loc[0,"Category"]
            a.price = str(tempDf.loc[0,"Price"])
            a.bonus = int(tempDf.loc[0,"Bonus"])
            a.limit = tempDf.loc[0,"Limit"]
            a.itemList0 = tempDf.drop_duplicates(subset='Name0')['Name0'].dropna(axis=0) + "[귀속] " + tempDf['Count0'].dropna(axis=0) + "개".splitlines()
            #a.itemList1 = tempDf.drop_duplicates(subset='Name1')["Name1"].dropna(axis=0) + "[귀속] " + tempDf['Count1'].dropna(axis=0) + "개 (" + tempDf['ItemID1'].dropna(axis=0) +")"
            a.itemList1 = tempDf.drop_duplicates(subset='Name1')["Name1"].dropna(axis=0) + "[귀속] " + tempDf['Count1'].dropna(axis=0) + "개"
            #a.itemList2 = tempDf.drop_duplicates(subset='Name2')["Name2"].dropna(axis=0) + "[귀속] " + str(tempDf['Count2'].dropna(axis=0)) + "개"
            a.server = tempDf.loc[0,"Server"]
            b = tempDf.drop_duplicates(subset='Name1')[["Name1","Name2"]].dropna(axis=0).index
            #b[0]=4, b[1]=6
            #print(a.itemName1)
            # for i in range(0, len(b)):
                

            #     a.itemName1[x] += tempDf.loc[0,"Name2"] + "[귀속] " + str(tempDf.loc[0,"Count2"]) + "개"

            if salesList != None :
                salesList.append(a)
            else :
                salesList = a


            #print(salesList[j].pkgName)

        #print(a.itemList2)
        del tempDf
        gc.collect()




    salesList.sort(key =lambda a: a.server)
    #for s in salesList:
    #    print(s.server)
    return salesList

def write_data_cashshop(salesList : list[Sales]):
    totalResult = pd.DataFrame()
#print(len(salesList))

    curRow = 0
    count = 0
    tqdmCount0=0
    print("데이터 쓰는 중...")
    for y in tqdm(salesList):
        tqdmCount0+=1
        y : Sales
        count += 1
        result = pd.DataFrame()

        i = curRow
        result.loc[i,"Category1"] = y.server
        result.loc[i,"Category2"] = y.pkgName + "\n" + str(y.pkgID)
        result.loc[i,"Category3"] = "이름"
        result.loc[i,"Check List"] = y.pkgName

    #■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        i += 1
        result.loc[i,"Category3"] = "카테고리"
        result.loc[i,"Check List"] = y.category
    #■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        i += 1
        result.loc[i,"Category3"] = "상세정보"
        desc0 = "\n".join(y.itemList0)
        desc0 = desc0.replace("다이아몬드[귀속]","다이아몬드")
        result.loc[i,"Check List"] = desc0

        i += 1
        desc1 = "\n".join(map(str, y.itemList1))
        desc1 = desc1.replace("nan\n","")
        desc1 = desc1.replace("\n","\n- ")
        result.loc[i,"Check List"] = "사용 시 다음 아이템 획득\n\n- "+desc1

        i += 1
        desc0 = desc0.replace("\n","\n- ")
        result.loc[i,"Check List"] = "<"+y.pkgName+"> 구성품 상세 정보\n- " + desc0

    #■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

        i += 1
        result.loc[i,"Category3"] = "상품슬롯"

        for item1 in y.itemList1 :
            if str(item1) != "nan" :
                result.loc[i,"Check List"] = str(item1) + " 관련 이미지 노출"
                i+=1

        result.loc[i,"Check List"] = "마일리지 : " + str(y.bonus)
        i+=1
        result.loc[i,"Check List"] = "구매 제한 : " + y.limit
        i+=1
        result.loc[i,"Check List"] = "구매 가격 : " + y.price
        
    #■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

        i += 1
        result.loc[i,"Category3"] = "아이템 구매"

        if "원" in y.price :
            result.loc[i,"Check List"] = f"결제 모듈 내 {y.pkgName} 노출"
            i += 1
            result.loc[i,"Check List"] = f"결제 모듈 내 {y.price} 노출"
            i += 1
            result.loc[i,"Check List"] = f"결제 완료 시 보관함으로 획득"
        else :
            result.loc[i,"Check List"] = y.price + " 차감"
    #■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

        i += 1
        result.loc[i,"Category3"] = "마일리지"

        if "0" in str(y.bonus) :
            result.loc[i,"Check List"] =  "마일리지 미노출"
        else :            
            result.loc[i,"Check List"] =  str(y.bonus)+ " 마일리지 적립"

    #■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        i += 1
        result.loc[i,"Category3"] = "아이템 획득"
        result.loc[i,"Check List"] = y.pkgName + "상자[귀속] 인벤토리 획득"
    #■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        i += 1
        result.loc[i,"Category3"] = "아이템 사용"

        for item1 in y.itemList1 :
            if str(item1) != "nan" :
                result.loc[i,"Check List"] = "- " + str(item1) + " 획득 및 사용"
                i+=1
        
        i-=1
    #■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        i += 1
        result.loc[i,"Category3"] = "구매 제한"
        result.loc[i,"Check List"] = y.limit + " 구매 시 상품 슬롯 비활성화"
        i += 1
        result.loc[i,"Check List"] = "상품 슬롯 하단에 [구매 완료] 라벨 노출"
    #■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        curRow = i
        

        totalResult = pd.concat([totalResult,result], ignore_index=True)
        #print(len(totalResult))

    totalResult = totalResult.replace("NaN","")
    totalResult = totalResult.replace("nan","")
    totalResult = totalResult.replace(np.nan,"")

    totalResult.to_excel(xlFileName, # directory and file name to write

                sheet_name = 'Sheet1', 

                na_rep = 'NaN', 

                float_format = "%.2f", 

                header = True, 

                #columns = ["group", "value_1", "value_2"], # if header is False

                index = True, 

                #index_label = "id", 

                startrow = 0, 

                startcol = 0, 

                #engine = 'xlsxwriter', 

                #freeze_panes = (2, 0)

                ) 


def postprocess_cashshop():
    wb = xl.load_workbook(xlFileName,data_only = True)
    sheetList = wb.sheetnames
    ws = wb[sheetList[0]]
    ws.column_dimensions['b'].width = 17
    ws.column_dimensions['c'].width = 24
    ws.column_dimensions['d'].width = 17
    ws.column_dimensions['e'].width = 60

    firstRow = 2
    lastRow = ws.max_row
    startRow_B =0
    startValue_B =""
    startRow_C = 0
    startRow_D = 0

    tqdmCount1 = 0
    print("엑셀 서식 처리중...")
    for i in tqdm(range(firstRow, lastRow+1)):
        tqdmCount1+=1
        #print(i)
        if (ws['b'+str(i)].value is not None) :
            if startRow_B == 0  :
                startRow_B = i
                startValue_B = ws['b'+str(i)].value
                #print(startRow_B)
            else :
                #firstTargetCell =  "C"+str(startRow_C)
                if ( ws['b'+str(i)].value != startValue_B) :
                    mergeTargetCell = "B"+str(startRow_B)+":B"+str(i-1)
                    ws.merge_cells(mergeTargetCell)
                    startValue_B = ws['b'+str(i)].value
                    startRow_B = i

        if ws['c'+str(i)].value is not None:
            if startRow_C == 0 :
                startRow_C = i
                #print(startRow)
            else :
                firstTargetCell =  "C"+str(startRow_C)
                mergeTargetCell = "C"+str(startRow_C)+":C"+str(i-1)
                ws.merge_cells(mergeTargetCell)
                startRow_C = i

        if ws['d'+str(i)].value is not None:
            if startRow_D == 0 :
                startRow_D = i
                #print(startRow)
            else :
                firstTargetCell =  "D"+str(startRow_D)
                mergeTargetCell = "D"+str(startRow_D)+":D"+str(i-1)
                ws.merge_cells(mergeTargetCell)
                startRow_D = i


        ws['b'+str(i)].alignment = Alignment(
            horizontal='center'
            ,vertical='top'
            ,wrap_text=True)
        ws['b'+str(i)].font = Font(size = 9, bold = True)
        ws['c'+str(i)].alignment = Alignment(
            horizontal='center'
            ,vertical='top'
            ,wrap_text=True)
        ws['c'+str(i)].font = Font(size = 9, bold = True)
        ws['d'+str(i)].alignment = Alignment(
            horizontal='center'
            ,vertical='top'
            ,wrap_text=True)
        ws['d'+str(i)].font = Font(size = 9, bold = True)
        ws['e'+str(i)].alignment = Alignment(
            horizontal='left'
            ,vertical='top'
            ,wrap_text=True)
        ws['e'+str(i)].font = Font(size = 9, bold = False)



    #예외 마지막 셀병합
    ws.merge_cells("B"+str(startRow_B)+":B"+str(lastRow))
    ws.merge_cells("C"+str(startRow_C)+":C"+str(lastRow))
    ws.merge_cells("D"+str(startRow_D)+":D"+str(lastRow))



    wb.save(xlFileName)


if __name__ == "__main__":

    #print("┃  R2M CASH SHOP CL MAKER  ┃")
    fileName = input("> 데이터파일명 입력(엔터:유료상점DATA.csv) : ")
    if fileName == "":
        fileName = "유료상점DATA.csv"

    while not os.path.isfile(fileName) :
        fileName = input("> Insert csv file name : ")


    salesList = extract_data_cashshop(fileName)
    write_data_cashshop(salesList)
    postprocess_cashshop()

    print("생성완료")
    input("종료하려면 엔터키 입력...")