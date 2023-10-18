'===========================================================
'Function to Create a Random Number with DateTime Stamp
'	Should modify this function to allow you to set the date as todays date rather than hardcoded to the date that the script was initially created
'===========================================================
Function fnRandomNumberWithDateTimeStamp()

'Find out the current date and time
Dim sDate : sDate = Day(Now)
Dim sMonth : sMonth = Month(Now)
Dim sYear : sYear = Year(Now)
Dim sHour : sHour = Hour(Now)
Dim sMinute : sMinute = Minute(Now)
Dim sSecond : sSecond = Second(Now)

'Create Random Number
fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)

End Function
'======================== End Function =====================

Dim OrderNumber

SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access").Maximize @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access").SAPGuiOKCode("OKCode").Set "/nva01" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access").SendKey ENTER @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("*Order Type").Set "OR" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf2.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Sales Organization").Set "1710" @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf2.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Distribution Channel").Set "10" @@ hightlight id_;_3_;_script infofile_;_ZIP::ssf2.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Division").Set "00" @@ hightlight id_;_4_;_script infofile_;_ZIP::ssf2.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Division").SetFocus @@ hightlight id_;_4_;_script infofile_;_ZIP::ssf2.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SendKey ENTER @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf2.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Sold-To Party").Set "EWM17-CU02" @@ hightlight id_;_3_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Ship-To Party").Set "EWM17-CU02" @@ hightlight id_;_4_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Cust. Reference").Set "450000019998" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Cust. Ref. Date").Set Month(Now) &"/" & Day(Now) & "/" & Year(Now)
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Item","10" @@ hightlight id_;_5_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Material","EWMS4-01" @@ hightlight id_;_5_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Order Quantity","1" @@ hightlight id_;_5_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Un","PC" @@ hightlight id_;_5_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SelectCell 1,"Un" @@ hightlight id_;_5_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SendKey F11 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf3.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiStatusBar("StatusBar").Sync @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf4.xml_;_

'===========================================================
'	Get the order number and store as a variable that could be used in the script, for example, if you wanted to do a va02 and/or a va03
'===========================================================
OrderNumber = SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiStatusBar("StatusBar").GetROProperty("item2") ' Output the order number as a variable
DataTable.Value("dtOrderNumber","Global") = OrderNumber

'msgbox OrderNumber

SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiButton("Exit").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf5.xml_;_

SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access").Maximize @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf6.xml_;_
SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access").SAPGuiButton("Exit").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf6.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Log Off").SAPGuiButton("Yes").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf7.xml_;_
