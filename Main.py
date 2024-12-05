import SchichtplanReader as spr
import SchichtplanWriter as spw
import SchichtplanStyler as sps

file_name = 'SCHICHTEINTEILUNG.xlsx'

dicList, sheetNames = spr.getData(file_name)
sheetNames.insert(0, "Name")
spw.writeData(dicList, file_name, sheetNames)
sps.styleSheet(file_name, list((dic["name"] for dic in dicList)), sheetNames)