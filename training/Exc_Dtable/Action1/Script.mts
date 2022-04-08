'................Data iterating through excel file with datatable..................
Option  explicit
Dim Rcount,File_Path,i


File_Path="C:\Users\sfjbs\Downloads\training\CExc01.xlsx"
datatable.ImportSheet File_Path,1,"Global"
datatable.GetSheet("Global")
Rcount=datatable.GetRowCount


For i = 1 To Rcount 
	datatable.SetCurrentRow(i)
	
Services.StartTransaction "Login"
	
	Systemutil.Run environment("App_Path")
	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set environment("Uname")
	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure "624d7f29bfba5671a353" @@ hightlight id_;_2135906456_;_script infofile_;_ZIP::ssf4.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click
	
Services.EndTransaction "Login"
	
	
	
Services.StartTransaction "Find_Flight"
	
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select datatable("From")
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("toCity").Select datatable("To") @@ hightlight id_;_2135911784_;_script infofile_;_ZIP::ssf9.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("Class").Select datatable("Class") @@ hightlight id_;_2134109208_;_script infofile_;_ZIP::ssf11.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTickets").Select datatable("Ticket") @@ hightlight id_;_2135882216_;_script infofile_;_ZIP::ssf13.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click
	
Services.EndTransaction "Find_Flight"
Reporter.ReportEvent micPass ,"Flight_Login" ,"successfully.............." 
	
Services.StartTransaction "Slct_Flight"
	
	WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0,1
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click
	
Services.EndTransaction "Slct_Flight"
Reporter.ReportEvent micPass, "Flight_Selection", "successfully.............." 
	
Services.StartTransaction "Order"

	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set  datatable("Pname")
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click
	
Services.EndTransaction "Order"
Reporter.ReportEvent micPass, "Flight_Order", "successfully.............." 	
	
	'WpfWindow("Micro Focus MyFlight Sample").WpfObject("WpfObject").Click -9,162
	WpfWindow("Micro Focus MyFlight Sample").Close @@ hightlight id_;_2819104_;_script infofile_;_ZIP::ssf21.xml_;_

	
Next

 @@ hightlight id_;_2135908232_;_script infofile_;_ZIP::ssf5.xml_;_

 @@ hightlight id_;_2134108296_;_script infofile_;_ZIP::ssf14.xml_;_

 @@ hightlight id_;_2135883032_;_script infofile_;_ZIP::ssf16.xml_;_

 @@ hightlight id_;_2135884616_;_script infofile_;_ZIP::ssf19.xml_;_



