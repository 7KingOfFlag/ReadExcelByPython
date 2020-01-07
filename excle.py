import xlrd
import io
import string

NumericList = ["int","short", "double", "float" ] # 数值列表
IntegeList = ["int","short"] # 整型列表
typeRow = 3 # 定义变量类型的行
propertysNameRow = 4 # 变量名所在的行
dataStartRow = 5 # 数据开始的行
exclePath = "./data.xlsx" # excle地址
dataSheetName = "data" # 数据表名

def DelimitationNumeric(value,typeName):
    """
    区分数值类型
    """
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
 
def AnalysisPropertys(dataSheet):
    """
    解析变量名
    """
    result = []
    for i in range(1,dataSheet.ncols):
        propertyType = dataSheet.cell(typeRow-1, i)
        propertysName = dataSheet.cell(propertysNameRow-1, i)
        result.append((propertysName,propertyType))
    return result

def GetPropertyName(propertyList,propertyIndex):
    """
    从列表中获取属性名
    """
    return propertyList[col - 1][0]

def GetPropertyType(propertyList,propertyIndex):
    """
    从列表中获取属性类型
    """
    return propertyList[col - 1][1]

if __name__ == '__main__':
    excle = xlrd.open_workbook(exclePath)
    dataSheet = excle.sheet_by_name(dataSheetName)
    propertyList = AnalysisPropertys(dataSheet)

    result = ''
    for row in range(dataStartRow - 1, dataSheet.nrows):
        result += '{'
        for col in range(1, dataSheet.ncols):
            context = dataSheet.cell(row,col)
            p = GetPropertyName(propertyList,col - 1)
            v = DelimitationNumeric(context,GetPropertyType())
            result += '\"{Name}\":{value}'.format(Name = p.value, value = v)
            if col != dataSheet.ncols -1:
                result += ','
        result += '}'
        if row != dataSheet.nrows - 1:
            result += '\n'
        pass
    print(result)