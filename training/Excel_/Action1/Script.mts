'.................iterating data in excel file............
Option  explicit
Dim i,j,Ccount,Rcount
Dim ExObj,WrkBk,sheet

Set ExObj=createobject("Excel.Application")
Set WrkBk=ExObj.Workbooks.Open("C:\Users\sfjbs\Downloads\training\Exl.xlsx")
Set sheet=WrkBk.Worksheets("Sheet1")

Ccount=sheet.usedrange.Columns.Count
Rcount=sheet.usedrange.Rows.Count

For i = 2 To Rcount
	For j = 1 To Ccount
		msgbox sheet.Cells(i,j)
	Next
Next
Set ExObj=nothing
Set WrkBk=nothing
Set sheet=nothing
