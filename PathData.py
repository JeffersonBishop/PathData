from urllib2 import urlopen
import requests
import xlsxwriter

try:
    urlFile = open("PathUrls.txt","r")
except:
    print("File Not Found")
    exit(-1)

fileText = urlFile.read()
fileText = fileText.split("\n")
Items = []
Prices = []


for line in fileText:
    if line == '':
        continue
    if line[0] == 'h':
        res = requests.get(line)
        res.raise_for_status()
        text = res.text
        text = text.split('>')
        for line in text:
            if line[0:24] == '<span data-tooltip title':
                lt = line.split('"')
                Prices.append(lt[1])
                break
    else:            
        Items.append(line)

xlsxFile = xlsxwriter.Workbook('PathItemData.xlsx')
spreadsheet = xlsxFile.add_worksheet()
x = 0
for i in Items:
    spreadsheet.write('A' + str(x+1), str(Items[x]))
    spreadsheet.write('B' + str(x+1), str(Prices[x]))
    x += 1
xlsxFile.close()


