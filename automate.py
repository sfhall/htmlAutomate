import xlrd

#structure of excel file must be: title, company, location-city, location-state
filename = input("Input the excel sheet:")
ss = xlrd.open_workbook(filename)
outputFile = open("placements.txt", "x") #creates new file

