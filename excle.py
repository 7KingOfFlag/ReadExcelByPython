import xlrd
import io
import string

NumericList = ["int","short", "double", "float" ] # 数值列表
IntegeList = ["int","short"] # 整型列表
#区分数值类型
def DelimitationNumeric(value,typeName):
    result = ""
    if typeName.value in NumericList:
        if typeName.value in IntegeList:
            result = '%d'%(value.value)
        else:
            result = value.value
    elif typeName.value == 'bool':
        # 布尔型
        result = value.value
    else:
        result = '\"{v}\"'.format(v = value.value)
    return result

excle = xlrd.open_workbook("./data.xlsx")
dataSheet = excle.sheet_by_name("data")

print('nRow:{nrow},nCol:{ncol}'.format(nrow = dataSheet.nrows, ncol = dataSheet.ncols))

typeRow = 3 # 定义变量类型的行
propertysNameRow = 4 # 变量名所在的行
dataStartRow = 5 # 数据开始的行

propertyList = []
propertysNameIndex = 0
propertyTypeIndex = 1
for i in range(1,dataSheet.ncols):
    print(i,typeRow-1)
    propertyType = dataSheet.cell(typeRow-1, i)
    propertysName = dataSheet.cell(propertysNameRow-1, i)
    propertyList.append((propertysName,propertyType))
    pass

result = '{'
for row in range(dataStartRow - 1, dataSheet.nrows):
    result += '{'
    for col in range(1, dataSheet.ncols):
        context = dataSheet.cell(row,col)
        p = propertyList[col - 1][propertysNameIndex]
        v = DelimitationNumeric(context,propertyList[col - 1][propertyTypeIndex])
        result += '\"{Name}\":{value}'.format(Name = p.value, value = v)
        if col != dataSheet.ncols -1:
            result += ','
    result += '}'
    if row != dataSheet.nrows - 1:
        result += ','
    pass
result += "}"
print(result)