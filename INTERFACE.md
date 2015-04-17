# js-xlsx

alaxls.write()
alaxls.writeFile()

# AlaSQL

ALAXLS()

workbook - format is the same as XLSX has

Two additions:
class = 
style = 


alasql('SELECT * INTO ALAXLSX("a.xlsx",?) FROM ?',[data,opt]);

h1range: A1
thrange: B3
tdrange: C4
thclass:string or function
tdclass :string or function(row,column)
thstyle: string or function
style: tdstyle: string or function

class:  
style: 

## alasql()

opts:
tableonly:true

SELECT a TH {width: expr, class: expr, style: expr } TD {class: expr, style: expr, format: expr, value: expr } AS '',

# Format Array

alaxls.apply(data, opt) => HTML

