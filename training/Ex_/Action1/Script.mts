'.....................data iterated through excel 1 and copied order details to excel 2...................
option explicit

Dim Rc,i,Fcity,Tcity,Psname,Fght_Path,File_Path,File02_Path,txt1
Dim Exc_Obj,Wrk_Bk,Sheet,F_Obj

Fght_Path="C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Micro Focus\Micro Focus UFT One\Sample Applications\Flight GUI"
File_Path="C:\Users\sfjbs\Downloads\training\Exc01.xlsx"
File02_Path="C:\Users\sfjbs\Downloads\training\Exc02.xlsx"

Set Exc_Obj=createobject("Excel.Application")
Set Wrk_Bk=Exc_Obj.Workbooks.Open(File_Path)
Set Sheet=Wrk_Bk.Worksheets("sheet1")
Set F_Obj=createobject("Scripting.FileSystemObject")

Rc=Sheet.usedrange.rows.count
For i = 2 To Rc
	Fcity=Sheet.Cells(i,"A")
	Tcity=Sheet.Cells(i,"B")
	Psname=Sheet.Cells(i,"C")
	
	Systemutil.Run Fght_Path
	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set "John"
	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure "624d04023ed7a89b09cd" @@ hightlight id_;_2118798184_;_script infofile_;_ZIP::ssf4.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click
	
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select Fcity
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("toCity").Select Tcity @@ hightlight id_;_2112162768_;_script infofile_;_ZIP::ssf9.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_2112164736_;_script infofile_;_ZIP::ssf10.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0,1 @@ hightlight id_;_2112164592_;_script infofile_;_ZIP::ssf11.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click @@ hightlight id_;_2118803128_;_script infofile_;_ZIP::ssf12.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set Psname @@ hightlight id_;_2112164256_;_script infofile_;_ZIP::ssf14.xml_;_
	
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click @@ hightlight id_;_2112166320_;_script infofile_;_ZIP::ssf15.xml_;_
	wait(3)
	
	If F_Obj.FileExists(File02_Path) Then
		Exc_Obj.Workbooks.Open(File02_Path)
		txt1=WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 87 completed").GetVisibleText
		Exc_Obj.Worksheets("sheet1").Cells(i,"A")=txt1
		Exc_Obj.ActiveWorkbook.Save
		
	else
		Exc_Obj.Workbooks.Add
		txt1=WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 87 completed").GetVisibleText
		Exc_Obj.Worksheets("sheet1").Cells(i,"A")=txt1
		Exc_Obj.ActiveWorkbook.SaveAs(File02_Path)
		
	End If
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 87 completed").Click 447,139 @@ hightlight id_;_2122131624_;_script infofile_;_ZIP::ssf16.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").Click 303,262
	WpfWindow("Micro Focus MyFlight Sample").Close
	
	
	
Next
Set Exc_Obj=nothing
Set Wrk_Bk=nothing
Set Sheet=nothing
 @@ hightlight id_;_2118799624_;_script infofile_;_ZIP::ssf5.xml_;_
 @@ hightlight id_;_6161812_;_script infofile_;_ZIP::ssf18.xml_;_
 @@ hightlight id_;_67970_;_script infofile_;_ZIP::ssf21.xml_;_
