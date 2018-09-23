#encoding: utf-8
#Jefferson Bishop, 2018
from urllib2 import urlopen
import requests
import xlsxwriter
import sys

def OpenFile():
    try:
        urlFile = open("PathUrls.txt","r")
        fileText = urlFile.read()
        fileText = fileText.split('\n')
        return fileText
    except:
        print("File Not Found")
        exit(-1)

def GetConversions():
    ConversionRates = {
            "Exalted": "http://currency.poe.trade/search?league=Standard&online=x&stock=&want=4&have=6",
            "Alchemy": "http://currency.poe.trade/search?league=Standard&online=x&stock=&want=4&have=3",
            "Alteration": "http://currency.poe.trade/search?league=Standard&online=x&stock=&want=4&have=1"
            }
    x = 0
    for line in ConversionRates.values():
        res = requests.get(line)
        res.raise_for_status()
        res.text.encode("utf-8")
        text = res.text.split('>')
        for line in text:
            if '&rarr' in line:
                lt = line.split(" ")
                break
        ConversionRates.update({ConversionRates.keys()[x]:float(lt[2])})
        x += 1
    return ConversionRates

def GetData(filetext, Rates):
    Prices = []
    Items = []
    Averages = []
    for line in filetext:
        if line == '':
            continue
        if line[0:4] == 'http':
            res = requests.get(line)
            res.raise_for_status()
            text = res.text
            text = text.split('>')
            x = 0
            All = []
            for line in text:
                if line[0:24] == '<span data-tooltip title':
                    lt = line.split('"')
                    if x == 0:
                        low = []
                        low.append(lt[1])
                        low = ConvertToChaos(Rates,low)
                        Prices.append(low[0])
                    x += 1
                    All.append(lt[1])
            if All == []:
                Prices.append("0 chaos")
                Averages.append(0)
            else:    
                All = ConvertToChaos(Rates, All) 
                Averages.append(GetAverage(All))
        else:
            Items.append(line)
    return Items, Prices, Averages

def ConvertToChaos(Rates, List):
    i = 0
    for item in List:
        item = item.split(' ')
        if str(item[1]) == 'chaos':
            List[i] = str(item[0]  + ' ' + item[1])
        elif str(item[1]) == 'exalted':
            converted = float(item[0]) * float(Rates["Exalted"])
            List[i] = str(converted) + ' chaos' 
        elif str(item[1]) == 'alchemy':
            converted = float(item[0]) * float(Rates["Alchemy"])
            List[i] = str(converted) + ' chaos'
        elif str(item[1]) == 'alteration':
            converted = float(item[0]) * float(Rates["Alteration"])
            List[i] = str(converted) + ' chaos'
        else:
            del List[i]
        i += 1
    return List

def GetAverage(All):
    num = 0
    total = 0
    for item in All:
        item = item.split(' ')
        num += 1
        total += float(item[0])
    Average = total / num
    return Average

def CreateFile(Items, Prices, Averages):
    xlsxFile = xlsxwriter.Workbook('PathItemData.xlsx')
    spreadsheet = xlsxFile.add_worksheet()
    formating = xlsxFile.add_format({'border':1})
    spreadsheet.set_column(0, 20, 50, formating)
    x = 1
    spreadsheet.write('A1','Name of Item')
    spreadsheet.write('B1','Lowest Price')
    spreadsheet.write('C1','Average Price in Chaos')
    spreadsheet.write('D1','Profit in Chaos')
    for i in Items:
        spreadsheet.write('A' + str(x+1), str(Items[x-1]))
        spreadsheet.write('B' + str(x+1), str(Prices[x-1]))
        spreadsheet.write('C' + str(x+1), str(Averages[x-1]))
        t = Prices[x-1].split(' ')
        spreadsheet.write('D' + str(x+1), str(Averages[x-1] - float(t[0])))
        if Prices[x-1] == "0 chaos":
            spreadsheet.write('B' + str(x+1), "No Listings")
            spreadsheet.write('C' + str(x+1), "No Listings")
            spreadsheet.write('D' + str(x+1), "No Listings")
        x += 1
    xlsxFile.close()    

def main():
    Rates = GetConversions()
    fileText = OpenFile()
    Items, Prices, Averages = GetData(fileText, Rates)
    CreateFile(Items, Prices, Averages)
    print("Spreadsheet file Generated")
    exit(0)

if __name__ == '__main__':
    main()
