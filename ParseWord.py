# author : chenboyan
# used for text ectracting

import docx
import os
import xlwt
import re

# get full text from a docx
def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

# get keyword from keyWord.txt
def getKeyWord():
    keyWordList = []
    if not os.path.exists('keyWord.txt'):
        return []
    with open('keyWord.txt', 'r') as f:
        lines = f.readlines()
        for line in lines:
            line = line.strip()
            keyWordList.append(line)
    return keyWordList


def extactMessage(fullText):
    fullText = fullText.replace('国家税务局稽查局\n', '!!!!!!\n')
    fullText = fullText.replace('地方税务局稽查局\n', '######\n').replace('地方税务局税务稽查处\n', '######\n')\
        .replace('地方税务局税务检查处\n', '######\n').replace('地方税务局稽查处\n', '######\n').replace('地方税务局稽查管理处\n', '######\n')
    fullText = replaceText(fullText)
    splitText = fullText.split('\n')

    for line in splitText:
        if '!!!!!!' in line:
            newline = '+$+' + line
            fullText = fullText.replace(line, newline)
        elif '######' in line:
            newline = '-￥-' + line
            fullText = fullText.replace(line, newline)
    provinceDcit = {}
    provinceText = fullText.split('+$')
    keyWordList = getKeyWord()

    for province in provinceText:
        nationDict = {}
        localDict = {}
        key = ''
        lines = province.split('\n')
        for line in lines:
            line = line.strip()
            if '!!!!!!' in line:
                key = line.replace('+', '').replace('!!!!!!', '')
        cutText = province.split('-￥-')
        for keyWord in keyWordList:
            paragraphText = cutText[0].split('\n[')
            for paragraph in paragraphText:
                if paragraph.startswith(keyWord):
                    content = paragraph.replace(keyWord+']', '')
                    nationDict[keyWord] = content.replace(' ', '')
        for keyWord in keyWordList:
            paragraphText = cutText[-1].split('\n[')
            for paragraph in paragraphText:
                if paragraph.startswith(keyWord):
                    content = paragraph.replace(keyWord+']', '')
                    localDict[keyWord] = content.replace(' ', '')
        taxList = [nationDict, localDict]
        if not key == '':
            provinceDcit[key] = taxList
    return provinceDcit


def writeExcel(dict):
    work_book = xlwt.Workbook(encoding='utf-8')
    headlineStyle = xlwt.XFStyle()
    headfont = xlwt.Font()
    # 字体基本设置
    headfont.name = u'华文仿宋'  # 在这里改标题字体
    headfont.color = 'black'
    headfont.height = 240  # 在这里改标题字号
    headfont.bold = True
    headlineStyle.font = headfont

    contentStyle = xlwt.XFStyle()
    font = xlwt.Font()
    # 字体基本设置
    font.name = u'华文仿宋'  # 在这里改正文字体
    font.color = 'black'
    font.height = 220  # 在这里改正文字号
    al = xlwt.Alignment()
    al.horz = 0x01  # 设置左端对齐
    al.vert = 0x00  # 设置水平居中
    al.wrap = 1
    contentStyle.alignment = al
    contentStyle.font = font

    for key in dict.keys():
        try:
            sheet = work_book.add_sheet(key)
        except:
            print('invalid sheet {sheet}'.format(sheet=key))
            continue
        nationDict, localDict = dict[key][0], dict[key][-1]
        i, j = 0, 0
        for nationKey in nationDict.keys():
            try:
                sheet.write(0, i, nationKey, headlineStyle)
                sheet.write(1, i, nationDict[nationKey], contentStyle)
                i = i + 1
            except:
                continue
        for localKey in localDict.keys():
            try:
                sheet.write(2, j, localKey, headlineStyle)
                sheet.write(3, j, localDict[localKey], contentStyle)
                j = j + 1
            except:
                continue
    k = 1
    while os.path.exists('表' + str(k) + '.xls'):
        k = k + 1
    work_book.save('表' + str(k) + '.xls')
    print('解析数据完毕，写入表 ' + str(k))


def makefinalExcel(dict):
    year = input("请您输入年份：")
    keyWordMap = {}
    with open('finalExcelKeyWord.txt', 'r') as f:
        content = f.readlines()
        for line in content:
            line = line.strip()
            key = line.split(r'::')[0]
            originalValues = line.split(r'::')[1]
            value = []
            if ',' in originalValues:
                originalValueList = originalValues.split(',')
                for originalValue in originalValueList:
                    value.append(originalValue)
            else:
                value.append(originalValues)
            per = line.split(r'::')[2]
            if not len(per) == 1:
                per = ''
            valuePerList = [value, per]
            keyWordMap[key] = valuePerList

    work_book = xlwt.Workbook(encoding='utf-8')
    sheet = work_book.add_sheet(year)
    sheet.write(0, 0, year)
    sheet.write(1, 0, '地区')

    k = 2
    for province in dict.keys():
        sheet.write(k, 0, province + '国税')
        sheet.write(k+1, 0, province + '地税')
        index = 1
        for key in keyWordMap.keys():
            if k == 2:
                sheet.write(1, index, key)
            valueList = keyWordMap[key]
            value = valueList[0]
            per = valueList[1]
            content = extractValue(dict, province, value, per)
            sheet.write(k, index, content[0])
            sheet.write(k+1, index, content[1])
            index = index + 1
        k = k + 2
    tableIndex = 1
    while os.path.exists('关键字查找表' + str(tableIndex) + '.xls'):
        tableIndex = tableIndex + 1
    work_book.save('关键字查找表' + str(tableIndex) + '.xls')
    print('解析数据完毕，写入关键字查找表 ' + str(tableIndex))


def extractValue(dict, province, values, per):
    if province not in dict.keys():
        return []

    nationDict, localDict = dict[province][0], dict[province][-1]
    nationStr, localStr = '', ''
    nationStrList, localStrList = [], []
    pattern = r',|/|;|\'|`|\?|"|\~|!&|\(|\)|\_|，|。|、|；|·|！|…|（|）'

    for key in nationDict.keys():
        nationStr = nationStr + nationDict[key]
        nationStrList = re.split(pattern, nationStr)
    for key in localDict.keys():
        localStr = localStr + localDict[key]
        localStrList = re.split(pattern, localStr)

    nationNumber, localNumber = '', ''

    for val in values:
        for sentence in nationStrList:
            if val in sentence:
                if per not in sentence:
                    continue
                nationNumber = re.findall(r'\d+', sentence)
                if nationNumber == []:
                    continue

                start = sentence.find(nationNumber[0])
                if "年" in sentence and len(nationNumber) > 1:
                    start = sentence.find(nationNumber[-1])
                if per in sentence:
                    end = sentence.rfind(per)
                    nationNumber = sentence[start:end+1]
                    break

    for val in values:
        for sentence in localStrList:
            if val in sentence:
                if per not in sentence:
                    continue
                localNumber = re.findall(r'\d+', sentence)
                if localNumber == []:
                    continue
                start = sentence.find(localNumber[0])
                if "年" in sentence and len(localNumber) > 1:
                    start = sentence.find(localNumber[-1])
                if per in sentence:
                    end = sentence.rfind(per)
                    localNumber = sentence[start:end]
                    break
                elif per not in sentence:
                    continue

    return [nationNumber, localNumber]


def replaceText(fullText):
    replaceDict = {}
    if not os.path.exists('replace.txt'):
        return fullText
    with open('replace.txt', 'r') as f:
        lines = f.readlines()
        for line in lines:
            line = line.strip()
            splitline = line.split(':')
            replaceDict[splitline[0]] = splitline[-1]

    for key in replaceDict.keys():
        try:
            fullText = fullText.replace(key, replaceDict[key])
        except:
            continue
    return fullText


if __name__ == "__main__":
    full_text = getText(r'/Users/chenboyan/PycharmProjects/ParseWord/文本/2009.docx')
    provinceDict = extactMessage(full_text)
    writeExcel(provinceDict)
    # makefinalExcel(provinceDict)

